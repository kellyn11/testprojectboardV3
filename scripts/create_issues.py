import os
import re
import subprocess
from pathlib import Path
from docx import Document

OWNER = os.environ["OWNER"]
REPO = os.environ["REPO"]

DOCX_PATH = Path("input/stories.docx")


def run_gh(args):
    result = subprocess.run(
        ["gh"] + args,
        capture_output=True,
        text=True,
        check=True
    )
    return result.stdout.strip()


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def extract_story_rows_from_docx():
    doc = Document(str(DOCX_PATH))

    rows_out = []

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        header = [clean_text(c.text) for c in rows[0].cells[:3]]

        if header != ["SN", "User Stories", "Acceptance Criteria"]:
            continue

        for row in rows[1:]:
            cells = row.cells
            if len(cells) < 3:
                continue

            sn = clean_text(cells[0].text)
            story = clean_text(cells[1].text)
            ac_raw = cells[2].text.strip()

            if not re.fullmatch(r"\d+", sn or ""):
                continue

            ac_lines = []
            for line in ac_raw.splitlines():
                line = clean_text(line)
                if line:
                    ac_lines.append(line)

            rows_out.append({
                "id": f"US{sn}",
                "story": story,
                "acceptance": ac_lines,
            })

    return rows_out


def issue_exists(story_id):
    output = run_gh([
        "issue", "list",
        "--repo", f"{OWNER}/{REPO}",
        "--search", f'"{story_id} -" in:title',
        "--json", "title"
    ])

    issues = eval(output)

    for issue in issues:
        if issue["title"].startswith(f"{story_id} -"):
            return True

    return False


def create_issue(row):
    title = f"{row['id']} - {row['story']}"

    body = f"""User Story:
{row['story']}
"""

    run_gh([
        "issue", "create",
        "--repo", f"{OWNER}/{REPO}",
        "--title", title,
        "--body", body
    ])

    print(f"Created {row['id']}")


def add_ac_comment(row):
    title_search = f'"{row["id"]} -" in:title'

    output = run_gh([
        "issue", "list",
        "--repo", f"{OWNER}/{REPO}",
        "--search", title_search,
        "--json", "number"
    ])

    issues = eval(output)
    if not issues:
        return

    issue_number = issues[0]["number"]

    lines = ["Acceptance Criteria", ""]

    for ac in row["acceptance"]:
        lines.append(f"- [ ] {ac}")

    comment = "\n".join(lines)

    run_gh([
        "issue", "comment", str(issue_number),
        "--repo", f"{OWNER}/{REPO}",
        "--body", comment
    ])

    print(f"Added AC to {row['id']}")


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError("stories.docx not found")

    rows = extract_story_rows_from_docx()

    for row in rows:
        if issue_exists(row["id"]):
            print(f"Skipping {row['id']} (already exists)")
            continue

        create_issue(row)
        add_ac_comment(row)


if __name__ == "__main__":
    main()
