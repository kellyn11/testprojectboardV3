# Pandoc GitHub Starter

This starter repo sets up **Pandoc inside GitHub Actions** so you can convert a Word document (`.docx`) into Markdown on GitHub first.

It is designed as the first step for a later pipeline like:

`Word (.docx) -> Markdown -> structured issues -> GitHub Project board`

## What this repo does

- stores your Word document in `input/`
- runs a GitHub Actions workflow manually
- converts `.docx` to Markdown using Pandoc
- saves the converted file into `output/`
- optionally commits the generated Markdown back into the repo

## Folder structure

```text
.
├─ .github/
│  └─ workflows/
│     └─ convert-docx-to-markdown.yml
├─ input/
│  └─ stories.docx
├─ output/
│  └─ stories.md
└─ README.md
```

## How to use

### 1. Create a GitHub repository
Create a new repository on GitHub, for example:

`user-stories-pandoc`

### 2. Upload these files
Upload everything from this starter repo into your new GitHub repository.

### 3. Add your Word document
Put your source Word file at:

`input/stories.docx`

If your Word file has a different name, update the workflow file accordingly.

### 4. Run the workflow
In GitHub:

- go to **Actions**
- open **Convert DOCX to Markdown**
- click **Run workflow**

The workflow will:

- install Pandoc using the official Pandoc setup action
- convert `input/stories.docx` into `output/stories.md`
- commit `output/stories.md` back into the repository if there are changes

## GitHub Actions permissions

If the workflow cannot push changes back to the repo, check this in GitHub:

- **Settings** -> **Actions** -> **General**
- ensure workflow permissions allow **Read and write permissions**

## Why this is useful

Pandoc is excellent for converting `.docx` into Markdown, but it does **not** create GitHub issues or Project items by itself. Pandoc supports converting Word docx into Markdown and many other formats. GitHub Projects can then bulk-add issues and pull requests after they exist. See the Pandoc manual and GitHub Projects documentation. citeturn932057search2turn932057search4turn932057search9

This repo gives you the **GitHub part first**, exactly as requested.

## Next step after this works

Once the Markdown output looks good, the next layer is:

- parse `output/stories.md`
- create GitHub issues
- add them into a GitHub Project board
- optionally set Status / Priority fields

## Notes

- Pandoc conversion quality depends on how clean the original Word document is.
- If your stories are in tables, you may later prefer CSV extraction for cleaner import.
- If your stories are in paragraphs/bullets, Markdown is a good intermediate step.
