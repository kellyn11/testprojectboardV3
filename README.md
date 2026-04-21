# Stories Sync Report

## Repo structure

- `input/stories.docx` = source user stories
- `scripts/create_issues.py` = create/update GitHub issues
- `scripts/export_status.py` = export status report
- `.github/workflows/create-issues.yml` = run issue sync
- `.github/workflows/export-status.yml` = run report export

## How to use

1. Upload/update `input/stories.docx`
2. Run **Sync Issues from DOCX**
3. Update issue status on the project board
4. Run **Export Status Report**
5. Download `output/status_report.txt`
