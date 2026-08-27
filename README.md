# Download Papers From List

A Codex skill for downloading papers from an existing Word, Markdown, or Excel paper list.

## Features

- Reads `.docx`, `.md`, and `.xlsx` paper lists.
- Selects all papers, red titles, bold titles, red-or-bold titles, or red-and-bold titles.
- Distinguishes paper links from common code-repository links.
- Follows a supplied landing page to a clearly exposed PDF link.
- Produces Markdown and CSV reports containing successes, failures, skips, and reasons.
- Rejects private-network URLs and verifies that retained files are PDFs.
- Does not bypass authentication, paywalls, or CAPTCHAs.

## Install

Clone this repository into the Codex skills directory:

```powershell
git clone https://github.com/hfnjupt/download-papers-from-list.git `
  "$env:USERPROFILE\.codex\skills\download-papers-from-list"
```

Restart Codex if the skill does not appear immediately.

## Usage

Ask Codex, for example:

> Use `$download-papers-from-list` to download only the red paper titles in this Word document and report failures.

The bundled helper can also be run directly:

```powershell
python scripts/download_papers_from_list.py "C:\path\papers.docx" `
  --output-dir "C:\path\downloaded" `
  --select red
```

Run with `--dry-run` first to inspect the selection without downloading.

## Selection modes

- `all`
- `red`
- `bold`
- `marked` (red or bold)
- `red-and-bold`

See [SKILL.md](SKILL.md) for the complete Codex workflow.
