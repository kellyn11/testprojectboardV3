import json
import os
import re
import subprocess
from pathlib import Path

from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm


REPO = os.environ["REPO"]
DOCX_PATH = Path("input/stories.docx")
TXT_OUTPUT = Path("output/status_report.txt")
PDF_OUTPUT = Path("output/status_report.pdf")


def run_gh(args):
    result = subprocess.run(
        ["gh"] + args,
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout.strip()


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def extract_story_rows_from_docx():
    doc = Document(str(DOCX_PATH))

    section_candidates = []
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        if re.match(r"^\d+\.\d+\s+", txt):
            section_candidates.append(txt)

    section_index = 0
    current_section = "General"
    rows_out = []

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        header = [clean_text(c.text) for c in rows[0].cells[:3]]
        if len(header) < 3:
            continue

        if header[0] != "SN" or header[1] != "User Stories" or header[2] != "Acceptance Criteria":
            continue

        if section_index < len(section_candidates):
            current_section = section_candidates[section_index]
            section_index += 1

        for row in rows[1:]:
            cells = row.cells
            if len(cells) < 3:
                continue

            sn = clean_text(cells[0].text)
            story = clean_text(cells[1].text)
            ac_raw = cells[2].text.strip()

            if not re.fullmatch(r"\d+", sn or ""):
                continue
            if not story:
                continue

            ac_lines = []
            for line in ac_raw.splitlines():
                line = line.strip()
                if line:
                    ac_lines.append(line)

            rows_out.append(
                {
                    "sn": int(sn),
                    "id": f"US{sn}",
                    "section": current_section,
                    "story": story,
                    "acceptance": ac_lines,
                }
            )

    return rows_out


def get_issue_status_map():
    out = run_gh(
        [
            "issue",
            "list",
            "--repo",
            REPO,
            "--state",
            "all",
            "--limit",
            "200",
            "--json",
            "title,state,labels",
        ]
    )

    issues = json.loads(out)
    status_map = {}

    for issue in issues:
        title = issue.get("title", "")
        m = re.match(r"^US(\d+)\b", title)
        if not m:
            continue

        sn = int(m.group(1))
        state = (issue.get("state") or "").lower()
        labels = [lbl.get("name", "").lower() for lbl in issue.get("labels", [])]

        if state == "closed":
            status = "Done"
        elif "in-progress" in labels:
            status = "In Progress"
        else:
            status = "Todo"

        status_map[sn] = status

    return status_map


def marker_for_status(status: str) -> str:
    s = (status or "").strip().lower()
    if s == "done":
        return "[X]"
    if s == "in progress":
        return "[-]"
    return "[ ]"


def format_acceptance(ac_lines):
    if not ac_lines:
        return ""
    return "\n".join(ac_lines)


def write_txt_report(rows, status_map):
    done_count = 0
    total_count = len(rows)

    # widths
    sn_w = 4
    st_w = 8
    us_w = 55
    ac_w = 40

    def wrap(text, width):
        words = text.split()
        if not words:
            return [""]
        lines = []
        current = words[0]
        for word in words[1:]:
            test = current + " " + word
            if len(test) <= width:
                current = test
            else:
                lines.append(current)
                current = word
        lines.append(current)
        return lines

    def row_border():
        return "+" + "-" * (sn_w + 2) + "+" + "-" * (st_w + 2) + "+" + "-" * (us_w + 2) + "+" + "-" * (ac_w + 2) + "+"

    lines = []
    lines.append("Project Functional Requirement Progress")
    lines.append("")
    lines.append("Legend:")
    lines.append("[X] Done")
    lines.append("[-] In Progress")
    lines.append("[ ] Todo")
    lines.append("")

    current_section = None

    for row in rows:
        if row["section"] != current_section:
            current_section = row["section"]
            lines.append(current_section)
            lines.append(row_border())
            lines.append(
                f"| {'SN'.ljust(sn_w)} | {'Status'.ljust(st_w)} | {'User Stories'.ljust(us_w)} | {'Acceptance Criteria'.ljust(ac_w)} |"
            )
            lines.append(row_border())

        status = status_map.get(row["sn"], "Todo")
        marker = marker_for_status(status)
        if marker == "[X]":
            done_count += 1

        story_lines = wrap(row["story"], us_w)
        ac_lines = []
        for ac in row["acceptance"]:
            ac_lines.extend(wrap(ac, ac_w))
        if not ac_lines:
            ac_lines = [""]

        max_lines = max(len(story_lines), len(ac_lines))

        for i in range(max_lines):
            sn = str(row["sn"]).ljust(sn_w) if i == 0 else " " * sn_w
            st = marker.ljust(st_w) if i == 0 else " " * st_w
            us = story_lines[i].ljust(us_w) if i < len(story_lines) else " " * us_w
            ac = ac_lines[i].ljust(ac_w) if i < len(ac_lines) else " " * ac_w
            lines.append(f"| {sn} | {st} | {us} | {ac} |")

        lines.append(row_border())

    lines.append("")
    lines.append(f"Completion: {done_count} / {total_count} Completed")
    progress = round((done_count / total_count) * 100) if total_count else 0
    lines.append(f"Progress: {progress}%")

    TXT_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    TXT_OUTPUT.write_text("\n".join(lines), encoding="utf-8")


def write_pdf_report(rows, status_map):
    PDF_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    normal_style = styles["BodyText"]
    normal_style.fontName = "Helvetica"
    normal_style.fontSize = 9
    normal_style.leading = 11

    section_style = styles["Heading3"]
    section_style.fontName = "Helvetica-Bold"
    section_style.fontSize = 11
    section_style.leading = 13

    doc = SimpleDocTemplate(
        str(PDF_OUTPUT),
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    elements = []

    done_count = 0
    total_count = len(rows)

    elements.append(Paragraph("Project Functional Requirement Progress", title_style))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("Legend: [X] Done &nbsp;&nbsp;&nbsp; [-] In Progress &nbsp;&nbsp;&nbsp; [ ] Todo", normal_style))
    elements.append(Spacer(1, 10))

    current_section = None

    for row in rows:
        if row["section"] != current_section:
            current_section = row["section"]
            elements.append(Paragraph(current_section, section_style))
            elements.append(Spacer(1, 4))

            table_data = [[
                Paragraph("<b>SN</b>", normal_style),
                Paragraph("<b>Status</b>", normal_style),
                Paragraph("<b>User Stories</b>", normal_style),
                Paragraph("<b>Acceptance Criteria</b>", normal_style),
            ]]

            section_rows = [r for r in rows if r["section"] == current_section]

            for r in section_rows:
                status = status_map.get(r["sn"], "Todo")
                marker = marker_for_status(status)
                if marker == "[X]":
                    done_count += 1

                ac_text = "<br/>".join(r["acceptance"]) if r["acceptance"] else ""

                table_data.append([
                    Paragraph(str(r["sn"]), normal_style),
                    Paragraph(marker, normal_style),
                    Paragraph(r["story"], normal_style),
                    Paragraph(ac_text, normal_style),
                ])

            tbl = Table(
                table_data,
                colWidths=[15 * mm, 22 * mm, 110 * mm, 110 * mm],
                repeatRows=1,
            )

            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LEADING", (0, 0), (-1, -1), 11),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))

            elements.append(tbl)
            elements.append(Spacer(1, 10))

    progress = round((done_count / total_count) * 100) if total_count else 0
    elements.append(Paragraph(f"<b>Completion:</b> {done_count} / {total_count} Completed", normal_style))
    elements.append(Paragraph(f"<b>Progress:</b> {progress}%", normal_style))

    doc.build(elements)


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError(f"Missing {DOCX_PATH}")

    rows = extract_story_rows_from_docx()
    rows.sort(key=lambda x: (x["section"], x["sn"]))

    status_map = get_issue_status_map()

    write_txt_report(rows, status_map)
    write_pdf_report(rows, status_map)

    print(f"Generated {TXT_OUTPUT}")
    print(f"Generated {PDF_OUTPUT}")


if __name__ == "__main__":
    main()
