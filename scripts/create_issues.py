import json
import os
import re
import subprocess
from pathlib import Path
from docx import Document

OWNER = os.environ["OWNER"]
REPO = os.environ["REPO"]
PROJECT_NUMBER = int(os.environ["PROJECT_NUMBER"])

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


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def normalize_status(status: str) -> str:
    value = clean_text(status).lower()
    return VALID_STATUSES.get(value, "Backlog")


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

        header = [clean_text(c.text) for c in rows[0].cells]

        required_headers = ["SN", "Status", "User Stories", "Acceptance Criteria"]
        if header[:4] != required_headers:
            continue

        if section_index < len(section_candidates):
            current_section = section_candidates[section_index]
            section_index += 1

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
                "sn": int(sn),
                "id": f"US{sn}",
                "section": current_section,
                "status": status,
                "story": story,
                "acceptance": ac_lines,
            })

    return rows_out


def find_existing_issue(story_id):
    search_query = f'repo:{OWNER}/{REPO} in:title "{story_id} -"'
    output = run_gh([
        "issue", "list",
        "--repo", f"{OWNER}/{REPO}",
        "--state", "all",
        "--search", search_query,
        "--json", "number,title,state"
    ])

    issues = json.loads(output)

    for issue in issues:
        if issue["title"].startswith(f"{story_id} -"):
            return issue

    return None


def create_or_update_issue(row):
    story_id = row["id"]
    title = f"{story_id} - {row['story']}"

    body = f"""Section: {row['section']}

User Story:
{row['story']}
"""

    existing_issue = find_existing_issue(story_id)

    if existing_issue:
        issue_number = existing_issue["number"]

        run_gh([
            "issue", "edit", str(issue_number),
            "--repo", f"{OWNER}/{REPO}",
            "--title", title,
            "--body", body
        ])

        if existing_issue["state"] == "CLOSED":
            run_gh([
                "issue", "reopen", str(issue_number),
                "--repo", f"{OWNER}/{REPO}"
            ])

        print(f"Updated issue {story_id}")
        return issue_number

    output = run_gh([
        "issue", "create",
        "--repo", f"{OWNER}/{REPO}",
        "--title", title,
        "--body", body,
        "--json", "number"
    ])

    issue = json.loads(output)
    issue_number = issue["number"]

    print(f"Created issue {story_id}")
    return issue_number


def build_ac_comment(row):
    lines = ["Acceptance Criteria", ""]

    for ac in row["acceptance"]:
        lines.append(f"- [ ] {ac}")

    return "\n".join(lines)


def get_issue_comments(issue_number):
    output = run_gh([
        "api",
        f"repos/{OWNER}/{REPO}/issues/{issue_number}/comments"
    ])

    return json.loads(output)


def create_or_update_ac_comment(issue_number, row):
    comments = get_issue_comments(issue_number)
    ac_body = build_ac_comment(row)

    existing_ac_comment = None

    for comment in comments:
        body = comment.get("body", "")
        if body.startswith("Acceptance Criteria"):
            existing_ac_comment = comment
            break

    if existing_ac_comment:
        comment_id = existing_ac_comment["id"]

        run_gh([
            "api",
            "--method", "PATCH",
            f"repos/{OWNER}/{REPO}/issues/comments/{comment_id}",
            "-f", f"body={ac_body}"
        ])

        print(f"Updated AC comment for issue #{issue_number}")
    else:
        run_gh([
            "issue", "comment", str(issue_number),
            "--repo", f"{OWNER}/{REPO}",
            "--body", ac_body
        ])

        print(f"Created AC comment for issue #{issue_number}")


def get_project_id_and_status_field():
    query = """
    query($owner:String!, $number:Int!) {
      user(login:$owner) {
        projectV2(number:$number) {
          id
          fields(first:50) {
            nodes {
              ... on ProjectV2SingleSelectField {
                id
                name
                options {
                  id
                  name
                }
              }
            }
          }
        }
      }
    }
    """

    output = run_gh([
        "api", "graphql",
        "-f", f"query={query}",
        "-F", f"owner={OWNER}",
        "-F", f"number={PROJECT_NUMBER}",
    ])

    data = json.loads(output)
    project = data["data"]["user"]["projectV2"]

    status_field = None

    for field in project["fields"]["nodes"]:
        if field and field.get("name") == "Status":
            status_field = field
            break

    if not status_field:
        raise RuntimeError("Status field not found in project board.")

    return project["id"], status_field


def get_issue_node_id(issue_number):
    output = run_gh([
        "api",
        f"repos/{OWNER}/{REPO}/issues/{issue_number}"
    ])

    issue = json.loads(output)
    return issue["node_id"]


def add_issue_to_project(project_id, issue_node_id):
    mutation = """
    mutation($projectId:ID!, $contentId:ID!) {
      addProjectV2ItemById(input: {
        projectId: $projectId,
        contentId: $contentId
      }) {
        item {
          id
        }
      }
    }
    """

    output = run_gh([
        "api", "graphql",
        "-f", f"query={mutation}",
        "-F", f"projectId={project_id}",
        "-F", f"contentId={issue_node_id}",
    ])

    data = json.loads(output)
    return data["data"]["addProjectV2ItemById"]["item"]["id"]


def update_project_status(project_id, item_id, status_field, status_name):
    option_id = None

    for option in status_field["options"]:
        if option["name"].lower() == status_name.lower():
            option_id = option["id"]
            break

    if not option_id:
        raise RuntimeError(f"Status option not found in board: {status_name}")

    mutation = """
    mutation($projectId:ID!, $itemId:ID!, $fieldId:ID!, $optionId:String!) {
      updateProjectV2ItemFieldValue(input: {
        projectId: $projectId,
        itemId: $itemId,
        fieldId: $fieldId,
        value: {
          singleSelectOptionId: $optionId
        }
      }) {
        projectV2Item {
          id
        }
      }
    }
    """

    run_gh([
        "api", "graphql",
        "-f", f"query={mutation}",
        "-F", f"projectId={project_id}",
        "-F", f"itemId={item_id}",
        "-F", f"fieldId={status_field['id']}",
        "-F", f"optionId={option_id}",
    ])


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError(f"Missing {DOCX_PATH}")

    rows = extract_story_rows_from_docx()
    rows.sort(key=lambda x: x["sn"])

    project_id, status_field = get_project_id_and_status_field()

    for row in rows:
        issue_number = create_or_update_issue(row)
        create_or_update_ac_comment(issue_number, row)

        issue_node_id = get_issue_node_id(issue_number)
        item_id = add_issue_to_project(project_id, issue_node_id)
        update_project_status(project_id, item_id, status_field, row["status"])

        print(f"Synced {row['id']} to project status: {row['status']}")


if __name__ == "__main__":
    main()
