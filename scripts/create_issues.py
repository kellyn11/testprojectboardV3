import json
import os
import re
import subprocess
import time
from pathlib import Path
from docx import Document

OWNER = os.environ["OWNER"]
REPO = os.environ["REPO"]
PROJECT_OWNER = os.environ["PROJECT_OWNER"]
PROJECT_NUMBER = int(os.environ["PROJECT_NUMBER"])
DEFAULT_STATUS = os.environ.get("DEFAULT_STATUS", "Backlog")

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


def generate_title_from_story(story):
    match = re.search(r"I want to (.*?)(?: so that|$)", story, re.IGNORECASE)
    title = match.group(1).strip() if match else story.strip()

    if not title:
        return "Untitled Story"

    return title[:1].upper() + title[1:]


def extract_story_rows_from_docx():
    doc = Document(str(DOCX_PATH))
    rows_out = []

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        header = [clean_text(c.text) for c in rows[0].cells[:2]]

        if header != ["User Stories", "Acceptance Criteria"]:
            continue

        for row in rows[1:]:
            cells = row.cells
            if len(cells) < 2:
                continue

            story = clean_text(cells[0].text)
            ac_raw = cells[1].text.strip()

            if not story:
                continue

            ac_lines = []
            for line in ac_raw.splitlines():
                line = clean_text(line)
                if line:
                    ac_lines.append(line)

            rows_out.append({
                "title": generate_title_from_story(story),
                "story": story,
                "acceptance": ac_lines,
            })

    return rows_out


def find_existing_issue_by_title(title):
    output = run_gh([
        "issue", "list",
        "--repo", f"{OWNER}/{REPO}",
        "--state", "all",
        "--search", title,
        "--json", "number,title,state"
    ])

    issues = json.loads(output)

    for issue in issues:
        if clean_text(issue["title"]) == clean_text(title):
            return issue

    return None


def create_or_update_issue(row):
    title = row["title"]

    body = f"""User Story:

{row["story"]}
"""

    existing = find_existing_issue_by_title(title)

    if existing:
        issue_number = existing["number"]

        run_gh([
            "issue", "edit", str(issue_number),
            "--repo", f"{OWNER}/{REPO}",
            "--title", title,
            "--body", body
        ])

        if existing["state"] == "CLOSED":
            run_gh([
                "issue", "reopen", str(issue_number),
                "--repo", f"{OWNER}/{REPO}"
            ])

        issue_data = json.loads(run_gh([
            "api",
            f"repos/{OWNER}/{REPO}/issues/{issue_number}"
        ]))

        print(f"Updated issue #{issue_number}: {title}")
        return issue_number, issue_data["node_id"]

    output = run_gh([
        "api",
        f"repos/{OWNER}/{REPO}/issues",
        "-f", f"title={title}",
        "-f", f"body={body}"
    ])

    issue = json.loads(output)

    print(f"Created issue #{issue['number']}: {title}")
    return issue["number"], issue["node_id"]


def build_ac_comment(row):
    lines = ["Acceptance Criteria", ""]

    for ac in row["acceptance"]:
        lines.append(f"- [ ] {ac}")

    return "\n".join(lines)


def create_or_update_ac_comment(issue_number, row):
    ac_body = build_ac_comment(row)

    comments = json.loads(run_gh([
        "api",
        f"repos/{OWNER}/{REPO}/issues/{issue_number}/comments"
    ]))

    existing_ac_comment = None

    for comment in comments:
        body = comment.get("body", "")
        if body.startswith("Acceptance Criteria"):
            existing_ac_comment = comment
            break

    if existing_ac_comment:
        run_gh([
            "api",
            "--method", "PATCH",
            f"repos/{OWNER}/{REPO}/issues/comments/{existing_ac_comment['id']}",
            "-f", f"body={ac_body}"
        ])
        print(f"Updated AC checkbox comment for issue #{issue_number}")
    else:
        run_gh([
            "api",
            f"repos/{OWNER}/{REPO}/issues/{issue_number}/comments",
            "-f", f"body={ac_body}"
        ])
        print(f"Created AC checkbox comment for issue #{issue_number}")


def get_project_info():
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
        raise RuntimeError("Status field not found in project.")

    return project["id"], status_field


def get_project_item_id_for_issue(project_id, issue_number):
    query = """
    query($projectId:ID!) {
      node(id:$projectId) {
        ... on ProjectV2 {
          items(first:100) {
            nodes {
              id
              content {
                ... on Issue {
                  number
                  title
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
        "-F", f"projectId={project_id}"
    ])

    data = json.loads(output)
    items = data["data"]["node"]["items"]["nodes"]

    for item in items:
        content = item.get("content")
        if content and content.get("number") == issue_number:
            return item["id"]

    return None


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
        raise RuntimeError(f"Status '{status_name}' not found. Available: {available}")

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


def ensure_issue_in_project_with_status(project_id, status_field, issue_number, issue_node_id):
    item_id = get_project_item_id_for_issue(project_id, issue_number)

    if not item_id:
        item_id = add_issue_to_project(project_id, issue_node_id)
        print(f"Added issue #{issue_number} to project")

        # Give GitHub a short moment before setting status
        time.sleep(2)
    else:
        print(f"Issue #{issue_number} already exists in project")

    for attempt in range(1, 4):
        try:
            update_project_status(project_id, item_id, status_field, DEFAULT_STATUS)
            return
        except Exception as error:
            print(f"Status update failed for issue #{issue_number}, attempt {attempt}/3")
            print(error)
            time.sleep(2)

    raise RuntimeError(f"Failed to set status for issue #{issue_number} after 3 attempts.")


def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError("Missing input/stories.docx")

    rows = extract_story_rows_from_docx()

    print(f"Found {len(rows)} user stories in DOCX")

    if not rows:
        raise RuntimeError(
            "No stories found. Header must be: User Stories | Acceptance Criteria"
        )

    project_id, status_field = get_project_info()

    imported_issues = []

    for index, row in enumerate(rows, start=1):
        print(f"Importing {index}/{len(rows)}: {row['title']}")

        issue_number, issue_node_id = create_or_update_issue(row)
        create_or_update_ac_comment(issue_number, row)

        ensure_issue_in_project_with_status(
            project_id,
            status_field,
            issue_number,
            issue_node_id
        )

        imported_issues.append(issue_number)

    print("Import completed.")
    print(f"Expected DOCX stories: {len(rows)}")
    print(f"Imported/repaired issues: {len(imported_issues)}")
    print(f"Issue numbers: {imported_issues}")


if __name__ == "__main__":
    main()
