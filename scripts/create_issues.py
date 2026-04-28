import json
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


def clean_text(text):
    return re.sub(r"\s+", " ", (text or "").strip())


def extract_story_rows_from_docx():
    doc = Document(str(DOCX_PATH))
    rows_out = []

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        header = [clean_text(c.text) for c in rows[0].cells[:3]]

        # EXPECTS:
        # SN | User Stories | Acceptance Criteria
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

            if not story:
                continue

            ac_lines = []
            for line in ac_raw.splitlines():
                line = clean_text(line)
                if line:
                    ac_lines.append(line)

            rows_out.append({
                "id": f"US{sn}",
                "story": story,
                "acceptance": ac_lines
            })

    return rows_out


def find_existing_issue(row):
    output = run_gh([
        "issue", "list",
        "--repo", f"{OWNER}/{REPO}",
        "--state", "all",
        "--search", row["id"],
        "--json", "number,title,state"
    ])

    issues = json.loads(output)

    for issue in issues:
        if issue["title"].startswith(f'{row["id"]} -'):
            return issue

    return None


def create_issue(row):
    title = f'{row["id"]} - {row["story"]}'

    body = f"""User Story:

{row["story"]}
"""

    output = run_gh([
        "issue", "create",
        "--repo", f"{OWNER}/{REPO}",
        "--title", title,
        "--body", body,
        "--json", "number"
    ])

    issue = json.loads(output)
    return issue["number"]


def add_ac_comment(issue_number, row):
    lines = [
        "Acceptance Criteria",
        ""
    ]

    for ac in row["acceptance"]:
        lines.append(f"- [ ] {ac}")

    comment_body = "\n".join(lines)

    run_gh([
        "issue", "comment", str(issue_number),
        "--repo", f"{OWNER}/{REPO}",
        "--body", comment_body
    ])


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError("Missing input/stories.docx")

    rows = extract_story_rows_from_docx()

    if not rows:
        raise RuntimeError(
            "No stories found. Make sure Word table header is: "
            "SN | User Stories | Acceptance Criteria"
        )

    for row in rows:
        existing = find_existing_issue(row)

        if existing:
            print(f'Skipping {row["id"]} because issue already exists.')
            continue

        issue_number = create_issue(row)
        add_ac_comment(issue_number, row)

        print(f'Imported {row["id"]} as issue #{issue_number}')


if __name__ == "__main__":
    main()
