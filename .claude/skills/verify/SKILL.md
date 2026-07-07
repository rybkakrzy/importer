---
name: verify
description: How to runtime-verify changes in this repo (GUI headless drive, backend converter harness, diagnostics CLI) without the full DB/GCS/Entra stack.
---

# Verify recipe — D2 ViewerEditor

## GUI (Angular editor) — headless Chrome + CDP, no backend needed

1. Backup + disable auth: set `"auth": { "enabled": false }` in
   `D2GuiViewerEditor/src/assets/configs/config.json` (**restore after!** — file is tracked).
2. `npx ng serve --port 4299` (background). `/editor` without `masterId` passes all
   guards when auth is disabled (documentAccessGuard skips, resourceGuard bypasses).
   Backend-unreachable banners are expected and harmless.
3. Drive with headless Chrome + raw CDP over Node ≥22 native `WebSocket`
   (no Playwright installed in this repo). Working driver template:
   scratchpad `verify/cdp-drive.mjs` pattern — launch
   `chrome.exe --remote-debugging-port=9333 --headless=new --user-data-dir=<tmp>`,
   fetch `/json`, connect WS, use `Runtime.evaluate` / `Input.*` / `Page.captureScreenshot`.
4. Dev build exposes `window.ng`: `ng.getComponent(document.querySelector('d2-wysiwyg-editor'))`
   gives the component instance (e.g. `setContent(...)`, `toggleFormattingMarks(true)`),
   then `ng.applyChanges(cmp)` to flush signals.

## Backend converters — package-boundary harness, no API/DB needed

`DocxToHtmlConverter` / `HtmlToDocxConverter` are plain classes:
`dotnet new console` in scratchpad + project reference to
`D2ApiViewerEditor/D2ViewerEditor.Infrastructure/D2ViewerEditor.Infrastructure.csproj`,
then `reader.Convert(stream)` / `writer.Convert(html, footer: …)` — same calls the
handlers make. For before/after evidence: `git stash push -- <converter file>`,
run, `git stash pop`, run again.

## Fixture DOCX without Word

`tools/docx-diagnostics/lib/__tests__/fixtures.mjs` exports `makeZip()` (zero-dep ZIP
writer). **Gotcha:** the OpenXml SDK requires `[Content_Types].xml` + `_rels/.rels`
in the package — the regex-based diagnostics tools don't, so a fixture that works
for `inspect-docx` may still open as EMPTY content in the .NET reader. Always include
both parts when the fixture is meant for the converter.

## Diagnostics CLI

`node tools/docx-diagnostics/inspect-docx.mjs <file> [--json]` (exit 0/1/2),
`compare-docx.mjs <src> <exported>` (exit 1 on lost/visual regressions).
Sidecar JSON reports land in CWD — run from a scratch dir.
