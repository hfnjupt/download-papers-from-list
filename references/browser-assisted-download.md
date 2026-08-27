# Authenticated browser downloads

Use this mode only for selected entries whose supplied publisher URL needs JavaScript, an existing browser session, or an institutional sign-in. Typical examples include IEEE Xplore article pages. The user's request to download through their institution authorizes using the signed-in session for the selected papers; it does not authorize credential extraction, access-control bypass, or unrelated browsing.

## Prepare the queue

Run the bundled parser with `--dry-run` and the requested selection rule. Use the generated CSV as the browser queue. Process only rows with `selected=true`; preserve rows marked `not selected`.

Do not run a title search or substitute a different copy unless the user separately asks for discovery. The supplied landing-page URL is the source of truth.

## Use the signed-in browser

Use the available browser-control skill and keep one authenticated browser session for the batch.

1. Open the supplied article URL.
2. Verify that the visible article title matches the queued title after ignoring punctuation, whitespace, and capitalization differences. If it does not match, record `title mismatch` and do not click a download control.
3. If institutional sign-in or SSO is required, hand control to the user for credentials, MFA, CAPTCHA, or organization selection. Resume only after the user says sign-in is complete.
4. Locate a visible control whose accessible name is `PDF`, `Download PDF`, `View PDF`, `Full Text PDF`, or a close publisher-specific equivalent. Prefer the control in the article header or full-text actions area.
5. For IEEE Xplore, select the red `PDF` control beside `Cite This`. Do not confuse it with `Full Text Views`, citation export, supplementary material, advertisements, or a PDF icon belonging to another article.
6. Activate the control and wait for either a browser download, a PDF viewer, or a new PDF tab. If a viewer opens, use its supported download action. Do not repeatedly click while navigation or a download is pending.
7. Confirm the resulting file begins with `%PDF-`, is non-empty, and belongs to the expected title. Rename it with the same safe-title convention as the bundled script, honoring the no-overwrite default.

Process entries sequentially unless the browser provides unambiguous per-download filenames and state. Reuse the authenticated session, but never read, export, log, or copy browser cookies, tokens, passwords, or institutional credentials.

## Stop conditions

Record a failure and continue to the next selected entry when:

- The page title does not match the queued paper.
- The PDF control is absent or disabled after the page finishes loading.
- The account lacks full-text entitlement.
- A CAPTCHA, MFA prompt, or login step still needs the user.
- The control opens a purchase page or requests payment.
- The downloaded payload is not a valid PDF.

Never bypass access controls or use credentials outside the user's existing authorized browser session.

## Report browser outcomes

Produce `browser_download_report.md` and `browser_download_report.csv` in the output directory. Include at least:

- title
- source URL
- browser-resolved URL when visible
- status: `success`, `failed`, `skipped`, or `needs user login`
- saved file
- reason

Merge counts with the direct-download report in the user-facing summary. A page opening successfully is not a successful download; only a retained and validated PDF counts as success.
