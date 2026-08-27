---
name: download-papers-from-list
description: Parse paper lists in Word (.docx), Markdown (.md), or Excel (.xlsx), select entries by formatting such as red titles or bold titles, download the associated PDFs, and produce success/failure reports. Use when the user supplies a literature list and asks to download all or only visually marked papers. Do not use for discovering papers that are absent from the supplied list.
---

# Download Papers From List

Download papers already named in a user-provided `.docx`, `.md`, or `.xlsx` list. Preserve the source list; write PDFs and reports to a separate output directory.

## Interpret the request

Resolve these from the user's wording and the input file:

- Input document and output directory
- Selection rule: `all`, `red`, `bold`, `marked` (red or bold), or `red-and-bold`
- Whether existing PDF filenames may be replaced; default to no overwrite
- Any title regex the user supplies in addition to formatting

Map requests such as “只下载标红的” to `red`, “只下载加粗的” to `bold`, and “下载所有突出显示的” to `marked`. If the input format cannot represent the requested formatting, stop and explain rather than silently downloading everything.

## Inspect before downloading

Use `scripts/download_papers_from_list.py` in `--dry-run` mode first. Review the generated plan/report and confirm that:

- Titles and download addresses are paired correctly.
- Code repositories, dataset URLs, author pages, and unrelated links are not treated as paper downloads.
- The selected count matches the requested formatting rule.
- Ambiguous or missing URLs are reported, not guessed from the title.

Do not broaden the task into a web literature search. A publisher or repository landing-page URL already present in the list may be followed to its PDF link; do not search for a different copy unless the user asks.

## Execute downloads

Run the same command without `--dry-run`. Prefer the bundled workspace Python runtime because it normally includes `python-docx` and `openpyxl`.

```powershell
python scripts/download_papers_from_list.py "C:\path\papers.docx" --output-dir "C:\path\downloaded" --select red
```

Useful options:

- `--select all|red|bold|marked|red-and-bold`
- `--sheet NAME` for one Excel sheet; otherwise inspect all visible sheets
- `--title-regex PATTERN` to combine a textual filter with formatting
- `--workers N`, `--timeout SECONDS`, and `--max-mb N`
- `--overwrite` only when the user authorizes replacing existing files
- `--report-prefix NAME` for deterministic report names

The helper accepts only public HTTP(S) targets, follows validated redirects, limits file size, and requires a PDF signature before keeping a file. Treat authentication, CAPTCHA, paywall, access-denied, and missing-link results as failures; do not attempt to bypass them.

## Deliver results

Always return:

- Download directory
- Counts for selected, successful, failed, skipped, and not selected
- Markdown report path
- CSV report path
- A concise list of failures with their recorded reasons

The reports are authoritative. Do not claim a paper was downloaded merely because a request returned HTTP 200; the script must have retained a valid PDF.
