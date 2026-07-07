# docx-diagnostics

Zero-dependency DOCX fidelity diagnostics for D2 ViewerEditor. Pure Node.js
(ESM), no `npm install` needed — a DOCX is read as a ZIP and inflated with the
built-in `zlib`, matching this repo's zero-dependency conversion stack. The
tools **never modify** the analysed files.

## inspect-docx

Dumps the structure the importer cares about (sections, page geometry,
headers/footers with text + fields, relationships, tables, drawings/shapes,
manual page breaks) and flags elements the importer is known to drop.

```bash
npm run inspect-docx -- ./samples/source.docx        # console report + JSON sidecar
node inspect-docx.mjs ./samples/source.docx --json    # raw JSON to stdout (machine mode)
```

Exit codes: `0` clean · `1` data-loss-risk elements present · `2` critical
(unreadable package or a referenced header/footer/image part is missing).

## compare-docx

Structural (not byte-for-byte) diff between a source and an exported DOCX —
Word emits semantically equivalent OOXML with different ordering/ids, so we
compare meaning. Each difference is classified `equal | equivalent |
changed-but-supported | changed-with-visual-impact | unsupported | lost |
added`.

```bash
npm run compare-docx -- ./samples/source.docx ./samples/exported.docx
```

Exit code `1` when any `lost` / `unsupported` / `changed-with-visual-impact`
difference is found, else `0` — usable as a round-trip regression gate in CI.

## Self-test

```bash
npm test        # node lib/__tests__/run.mjs — builds synthesized DOCX fixtures
```

The fixtures are generated in-memory (`lib/__tests__/fixtures.mjs`, includes a
minimal ZIP writer); no real customer documents are committed.
