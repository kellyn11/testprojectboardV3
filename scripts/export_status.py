import json
import re
import subprocess
from pathlib import Path

from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
import os


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


def wrap_text(text: str, font_name: str, font_size: int, max_width: float):
    words = text.split()
    if not words:
        return [""]

    lines = []
    current = words[0]

    for word in words[1:]:
        test = current + " " + word
        if stringWidth(test, font_name, font_size) <= max_width:
            current = test
        else:
            lines.append(current)
            current = word

    lines.append(current)
    return lines


def write_txt_report(rows, status_map):
    done_count = 0
    total_count = len(rows)

    lines = [
        "Project Functional Requirement Progress",
        "",
        "Legend:",
        "[X] Done",
        "[-] In Progress",
        "[ ] Todo",
        "",
    ]

    current_section = None
    for row in rows:
        if row["section"] != current_section:
            current_section = row["section"]
            lines.append(current_section)

        status = status_map.get(row["sn"], "Todo")
        marker = marker_for_status(status)
        if marker == "[X]":
            done_count += 1

        lines.append(f"{marker} {row['id']} - {row['story']}")
        lines.append("")

    lines.append(f"Completion: {done_count} / {total_count} Completed")
    progress = round((done_count / total_count) * 100) if total_count else 0
    lines.append(f"Progress: {progress}%")

    TXT_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    TXT_OUTPUT.write_text("\n".join(lines), encoding="utf-8")


def write_pdf_report(rows, status_map):
    PDF_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    c = canvas.Canvas(str(PDF_OUTPUT), pagesize=A4)
    width, height = A4

    left = 50
    right = 50
    top = height - 50
    bottom = 50
    line_gap = 16
    usable_width = width - left - right

    done_count = sum(
        1 for r in rows if marker_for_status(status_map.get(r["sn"], "Todo")) == "[X]"
    )
    total_count = len(rows)
    progress = round((done_count / total_count) * 100) if total_count else 0

    y = top

    def new_page():
        nonlocal y
        c.showPage()
        y = top

    def draw_line(text, font="Helvetica", size=11, indent=0):
        nonlocal y
        max_width = usable_width - indent
        wrapped = wrap_text(text, font, size, max_width)
        for part in wrapped:
            if y < bottom:
                new_page()
            c.setFont(font, size)
            c.drawString(left + indent, y, part)
            y -= line_gap

    c.setTitle("Project Functional Requirement Progress")

    draw_line("Project Functional Requirement Progress", font="Helvetica-Bold", size=16)
    y -= 4
    draw_line("Legend:", font="Helvetica-Bold", size=11)
    draw_line("[X] Done")
    draw_line("[-] In Progress")
    draw_line("[ ] Todo")
    y -= 6

    current_section = None
    for row in rows:
        section = row["section"]
        if section != current_section:
            current_section = section
            y -= 4
            draw_line(section, font="Helvetica-Bold", size=12)

        status = status_map.get(row["sn"], "Todo")
        marker = marker_for_status(status)
        draw_line(f"{marker} {row['id']} - {row['story']}", font="Helvetica", size=11)
        y -= 2

    y -= 8
    draw_line(f"Completion: {done_count} / {total_count} Completed", font="Helvetica-Bold", size=11)
    draw_line(f"Progress: {progress}%", font="Helvetica-Bold", size=11)

    c.save()


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError(f"Missing {DOCX_PATH}")

    rows = extract_story_rows_from_docx()
    rows.sort(key=lambda x: x["sn"])

    status_map = get_issue_status_map()

    write_txt_report(rows, status_map)
    write_pdf_report(rows, status_map)

    print(f"Generated {TXT_OUTPUT}")
    print(f"Generated {PDF_OUTPUT}")


if __name__ == "__main__":
    main()
