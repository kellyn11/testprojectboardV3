import json
import os
import re
import subprocess
from pathlib import Path
from docx import Document

OWNER = os.environ["OWNER"]
REPO = os.environ["REPO"]
DOCX_PATH = Path("input/stories.docx")

VALID_STATUSES = {
    "backlog": "Backlog",
    "todo": "Todo",
    "to-do": "Todo",
    "in progress": "In Progress",
    "qa review": "QA Review",
    "done": "Done",
}


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


def normalize_status(status):
    value = clean_text(status).lower()
    return VALID_STATUSES.get(value, "Backlog")


def extract_story_rows_from_docx():
    doc = Document(str(DOCX_PATH))
    rows_out = []

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        header = [clean_text(c.text) for c in rows[0].cells[:4]]

        if header != ["SN", "Status", "User Stories", "Acceptance Criteria"]:
            continue

        for row in rows[1:]:
            cells = row.cells
            if len(cells) < 4:
                continue

            sn = clean_text(cells[0].text)
            status = normalize_status(cells[1].text)
            story = clean_text(cells[2].text)
            ac_raw = cells[3].text.strip()

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
                "status": status,
                "story": story,
                "acceptance": ac_lines,
            })

    return rows_out


def find_existing_issue(row):
    output = run_gh([
        "issue", "list",
        "--repo", f"{OWNER}/{REPO}",
        "--state", "all",
        "--search", f'{row["id"]} in:title',
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

Imported Status:
{row["status"]}
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
            "No stories found. Make sure your Word table header is: "
            "SN | Status | User Stories | Acceptance Criteria"
        )

    for row in rows:
        existing = find_existing_issue(row)

        if existing:
            print(f'Skipping {row["id"]} because issue already exists.')
            continue

        issue_number = create_issue(row)
        add_ac_comment(issue_number, row)

        print(f'Imported {row["id"]} as issue #{issue_number} with status {row["status"]}')


if __name__ == "__main__":
    main()
