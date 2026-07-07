#!/usr/bin/env node
// inspect-docx — dump the structure of a DOCX package for fidelity diagnostics.
//
//   node inspect-docx.mjs <file.docx> [--json]
//
//   default : human-readable console report + a `<name>.report.json` sidecar.
//   --json  : print the raw JSON report to stdout (machine mode), no sidecar.
//
// Exit codes: 0 = clean, 1 = data-loss-risk elements present,
//             2 = critical (unreadable package / missing referenced part).
// The tool never modifies the analysed file.

import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { readZip } from './lib/zip.mjs';
import { buildReport } from './lib/docx-report.mjs';

function main(argv) {
  const args = argv.slice(2);
  const jsonMode = args.includes('--json');
  const file = args.find((a) => !a.startsWith('--'));
  if (!file) {
    console.error('Usage: node inspect-docx.mjs <file.docx> [--json]');
    return 2;
  }

  let report;
  try {
    const buf = readFileSync(file);
    report = buildReport(readZip(buf));
  } catch (e) {
    console.error(`inspect-docx: cannot read "${file}": ${e.message}`);
    return 2;
  }

  const exit = severityExitCode(report);

  if (jsonMode) {
    process.stdout.write(JSON.stringify(report, null, 2) + '\n');
    return exit;
  }

  printConsole(file, report);
  const outPath = `${basename(file).replace(/\.docx$/i, '')}.report.json`;
  writeFileSync(outPath, JSON.stringify(report, null, 2));
  console.log(`\nJSON report written to ${outPath}`);
  return exit;
}

function severityExitCode(report) {
  if (report.warnings.some((w) => w.severity === 'critical')) return 2;
  if (report.unsupportedElements.length > 0) return 1;
  return 0;
}

function printConsole(file, r) {
  const h = (t) => console.log(`\n\x1b[1m${t}\x1b[0m`);
  console.log(`\x1b[1mDOCX inspection:\x1b[0m ${file}`);

  h(`Sections (${r.sections.length})`);
  r.pageSettings.forEach((ps) => {
    console.log(
      `  #${ps.sectionIndex} ${ps.orientation}  ` +
        `page ${ps.physicalPageWidth.cm}×${ps.physicalPageHeight.cm} cm  ` +
        `margins T${ps.margins.top.cm}/R${ps.margins.right.cm}/B${ps.margins.bottom.cm}/L${ps.margins.left.cm} cm  ` +
        `hdr ${ps.margins.header.cm} ftr ${ps.margins.footer.cm}  ` +
        `content ${ps.contentAreaWidth.cm}×${ps.contentAreaHeight.cm} cm`,
    );
  });

  h(`Footers (${r.footers.length}) / Headers (${r.headers.length})`);
  for (const f of [...r.footers, ...r.headers]) {
    console.log(`  ${f.file} [${f.type}] text="${truncate(f.text, 60)}" fields=[${f.fields.join(',')}] tbl=${f.tableCount} img=${f.imageCount} shapes=${f.shapeCount}`);
    for (const rel of f.relationships) if (rel.missing) console.log(`    \x1b[31mMISSING image target: ${rel.target}\x1b[0m`);
  }

  h('Footer references');
  for (const ref of r.footerReferences) {
    const flag = ref.missing ? '\x1b[31m MISSING\x1b[0m' : '';
    console.log(`  section ${ref.sectionIndex}: ${ref.type} → ${ref.relationshipId} (${ref.xmlFile ?? '—'})${flag}`);
  }

  h(`Fields (${r.fields.length})`);
  for (const f of tally(r.fields.map((f) => f.kind))) console.log(`  ${f.key}: ${f.count}`);
  const dyn = r.fields.filter((f) => ['page-number', 'page-count', 'section-page-count'].includes(f.kind));
  if (dyn.length) console.log(`  → ${dyn.length} page-number/count field(s) — must round-trip as w:fld, not frozen text.`);

  h(`Tables (${r.tables.length})`);
  for (const t of r.tables) console.log(`  #${t.index} rows=${t.rows} cols=${t.gridColumns} gridSpan=${t.gridSpan} vMerge=${t.vMerge} styled=${t.styled} nested=${t.nested}`);

  h(`Drawings (${r.drawings.length}) / Shapes (${r.shapes.length})`);
  for (const d of r.drawings) console.log(`  drawing #${d.index} ${d.mode} img=${d.hasImage} chart=${d.hasChart} textbox=${d.hasTextbox} wrap=${d.wrap}`);
  for (const s of r.shapes) console.log(`  shape [${s.format}] textbox=${s.hasTextbox} table=${s.hasTable}`);

  console.log(`\nManual page breaks: ${r.manualPageBreaks}`);
  console.log(`Relationships: ${r.relationships.length}`);

  if (r.unsupportedElements.length) {
    h(`\x1b[33mUnsupported elements — DATA-LOSS RISK (${r.unsupportedElements.length})\x1b[0m`);
    for (const u of r.unsupportedElements) console.log(`  \x1b[33m${u.element} ×${u.count} in ${u.part} — ${u.note}\x1b[0m`);
  }
  if (r.warnings.length) {
    h(`\x1b[31mWarnings (${r.warnings.length})\x1b[0m`);
    for (const w of r.warnings) console.log(`  [${w.severity}] ${w.path}: ${w.message}`);
  }
}

function tally(arr) {
  const m = new Map();
  for (const k of arr) m.set(k, (m.get(k) || 0) + 1);
  return [...m.entries()].map(([key, count]) => ({ key, count }));
}
function truncate(s, n) { return (s || '').length > n ? s.slice(0, n) + '…' : s || ''; }

process.exit(main(process.argv));
