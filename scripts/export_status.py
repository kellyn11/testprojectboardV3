import os
import subprocess

REPO = os.environ["REPO"]
OUTPUT = "output/status_report.txt"


def get_issues():
    result = subprocess.run(
        [
            "gh", "issue", "list",
            "--repo", REPO,
            "--state", "all",
            "--json", "title,state",
        ],
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout


with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write("Project Functional Requirement Progress\n\n")
    f.write("Legend:\n")
    f.write("[X] Done\n")
    f.write("[-] In Progress\n")
    f.write("[ ] Todo\n\n")
    f.write("Export stub created.\n")
