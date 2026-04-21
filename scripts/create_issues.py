import os
import re
import subprocess
from docx import Document

REPO = os.environ["REPO"]
DOCX_PATH = "input/stories.docx"


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def make_title(sn: str, story: str) -> str:
    text = clean_text(story)
    text = re.sub(r"^As an? .*?,\s*I want to\s*", "", text, flags=re.I)
    text = re.sub(r"\s+so that.*$", "", text, flags=re.I)
    text = text.strip().rstrip(".")
    if not text:
        text = f"Story {sn}"
    return f"US{sn} - {text[:80]}"


def find_existing_issue(sn: str):
    result = subprocess.run(
        [
            "gh", "issue", "list",
            "--repo", REPO,
            "--state", "all",
            "--search", f"\"US{sn} -\" in:title",
            "--json", "number,title",
            "--jq", f'.[] | select(.title | startswith("US{sn} -")) | [.number, .title] | @tsv'
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    out = result.stdout.strip()
    if not out:
        return None, None
    number, title = out.split("\t", 1)
    return number, title


def create_issue(title: str, body: str):
    subprocess.run(
        [
            "gh", "issue", "create",
            "--repo", REPO,
            "--title", title,
            "--body", body,
        ],
        check=True,
    )


def update_issue(issue_number: str, title: str, body: str):
    subprocess.run(
        [
            "gh", "issue", "edit", issue_number,
            "--repo", REPO,
            "--title", title,
            "--body", body,
        ],
        check=True,
    )


def main():
    doc = Document(DOCX_PATH)

    current_section = "General"
    created = 0
    updated = 0

    section_candidates = []
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        if re.match(r"^\d+\.\d+\s+", txt):
            section_candidates.append(txt)

    section_index = 0

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

            title = make_title(sn, story)

            ac_lines = []
            for line in ac_raw.splitlines():
                line = line.strip()
                if line:
                    ac_lines.append(line)

            body_parts = [
                f"**Section:** {current_section}",
                "",
                "**User Story**",
                story,
                "",
                "**Acceptance Criteria**",
            ]

            if ac_lines:
                body_parts.extend(ac_lines)
            elif ac_raw.strip():
                body_parts.append(ac_raw.strip())
            else:
                body_parts.append("No acceptance criteria provided.")

            body = "\n".join(body_parts).strip()

            issue_number, old_title = find_existing_issue(sn)

            if issue_number:
                print(f"Updating issue #{issue_number}: {old_title} -> {title}")
                update_issue(issue_number, title, body)
                updated += 1
            else:
                print(f"Creating issue: {title}")
                create_issue(title, body)
                created += 1

    print(f"Done. Created {created} issue(s), updated {updated} issue(s).")


if __name__ == "__main__":
    main()
