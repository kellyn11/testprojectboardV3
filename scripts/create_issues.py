import json
import os
import re
import subprocess
from pathlib import Path
from docx import Document

OWNER = os.environ["OWNER"]
REPO = os.environ["REPO"]
PROJECT_OWNER = os.environ["PROJECT_OWNER"]
PROJECT_NUMBER = int(os.environ["PROJECT_NUMBER"])
DEFAULT_STATUS = os.environ.get("DEFAULT_STATUS", "Todo")

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
        "api",
        f"repos/{OWNER}/{REPO}/issues",
        "-f", f"title={title}",
        "-f", f"body={body}"
    ])

    issue = json.loads(output)

    print(f'Created {row["id"]} as issue #{issue["number"]}')
    return issue["number"], issue["node_id"]


def add_ac_comment(issue_number, row):
    lines = ["Acceptance Criteria", ""]

    for ac in row["acceptance"]:
        lines.append(f"- [ ] {ac}")

    comment_body = "\n".join(lines)

    run_gh([
        "api",
        f"repos/{OWNER}/{REPO}/issues/{issue_number}/comments",
        "-f", f"body={comment_body}"
    ])

    print(f'Added AC checklist comment to issue #{issue_number}')


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
        "-F", f"owner={PROJECT_OWNER}",
        "-F", f"number={PROJECT_NUMBER}"
    ])

    data = json.loads(output)
    project = data["data"]["user"]["projectV2"]

    if not project:
        raise RuntimeError("Project not found. Check PROJECT_OWNER and PROJECT_NUMBER.")

    status_field = None

    for field in project["fields"]["nodes"]:
        if field and field.get("name") == "Status":
            status_field = field
            break

    if not status_field:
        raise RuntimeError("Status field not found in project board.")

    return project["id"], status_field


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
        "-F", f"contentId={issue_node_id}"
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
        available = [option["name"] for option in status_field["options"]]
        raise RuntimeError(
            f"Status option '{status_name}' not found. Available: {available}"
        )

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
        "-F", f"optionId={option_id}"
    ])

    print(f"Moved project item to {status_name}")


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError("Missing input/stories.docx")

    rows = extract_story_rows_from_docx()

    if not rows:
        raise RuntimeError(
            "No stories found. Make sure Word table header is: "
            "SN | User Stories | Acceptance Criteria"
        )

    project_id, status_field = get_project_id_and_status_field()

    for row in rows:
        existing = find_existing_issue(row)

        if existing:
            print(f'Skipping {row["id"]} because issue already exists.')
            continue

        issue_number, issue_node_id = create_issue(row)
        add_ac_comment(issue_number, row)

        item_id = add_issue_to_project(project_id, issue_node_id)
        update_project_status(project_id, item_id, status_field, DEFAULT_STATUS)

        print(f'Imported {row["id"]} into project board as {DEFAULT_STATUS}')


if __name__ == "__main__":
    main()
