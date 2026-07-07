#!/usr/bin/env node
// compare-docx — structural diff between a source DOCX and an exported DOCX to
// catch fidelity/round-trip regressions.
//
//   node compare-docx.mjs <source.docx> <exported.docx> [--json]
//
// Compares semantics (page geometry, headers/footers, relationships, fields,
// tables, drawings/shapes, manual page breaks), NOT raw XML bytes — Word may
// emit semantically equivalent OOXML with different ordering/ids.
//
// Each difference is classified:
//   equal | equivalent | changed-but-supported |
//   changed-with-visual-impact | unsupported | lost | added
//
// Exit codes: 0 = no material regressions, 1 = data-loss / visual regressions.

import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { readZip } from './lib/zip.mjs';
import { buildReport } from './lib/docx-report.mjs';

const REGRESSION = new Set(['lost', 'unsupported', 'changed-with-visual-impact']);

function main(argv) {
  const args = argv.slice(2);
  const jsonMode = args.includes('--json');
  const files = args.filter((a) => !a.startsWith('--'));
  if (files.length < 2) {
    console.error('Usage: node compare-docx.mjs <source.docx> <exported.docx> [--json]');
    return 2;
  }

  let a, b;
  try {
    a = buildReport(readZip(readFileSync(files[0])));
    b = buildReport(readZip(readFileSync(files[1])));
  } catch (e) {
    console.error(`compare-docx: ${e.message}`);
    return 2;
  }

  const diffs = compare(a, b);
  const regressions = diffs.filter((d) => REGRESSION.has(d.status));
  const exit = regressions.length ? 1 : 0;

  if (jsonMode) {
    process.stdout.write(JSON.stringify({ source: files[0], exported: files[1], diffs }, null, 2) + '\n');
    return exit;
  }

  printConsole(files, diffs, regressions);
  const outPath = `${basename(files[0]).replace(/\.docx$/i, '')}-vs-${basename(files[1]).replace(/\.docx$/i, '')}.compare.json`;
  writeFileSync(outPath, JSON.stringify({ source: files[0], exported: files[1], diffs }, null, 2));
  console.log(`\nJSON diff written to ${outPath}`);
  return exit;
}

function compare(a, b) {
  const diffs = [];
  const push = (aspect, status, from, to, note) => diffs.push({ aspect, status, from, to, note });

  // Section count.
  if (a.sections.length !== b.sections.length) {
    push('sections.count', b.sections.length < a.sections.length ? 'lost' : 'added', a.sections.length, b.sections.length, 'Section count changed (multi-section flattening is a known risk).');
  }

  // Page geometry per section.
  const n = Math.max(a.pageSettings.length, b.pageSettings.length);
  for (let i = 0; i < n; i++) {
    const pa = a.pageSettings[i];
    const pb = b.pageSettings[i];
    if (!pa || !pb) { push(`section[${i}]`, pa ? 'lost' : 'added', !!pa, !!pb, 'Section present in only one document.'); continue; }
    geomDiff(push, i, 'page.width', pa.physicalPageWidth.twips, pb.physicalPageWidth.twips);
    geomDiff(push, i, 'page.height', pa.physicalPageHeight.twips, pb.physicalPageHeight.twips);
    if (pa.orientation !== pb.orientation) push(`section[${i}].orientation`, 'changed-with-visual-impact', pa.orientation, pb.orientation, '');
    for (const side of ['top', 'right', 'bottom', 'left', 'header', 'footer']) {
      geomDiff(push, i, `margin.${side}`, pa.margins[side].twips, pb.margins[side].twips);
    }
  }

  // Header/footer text presence.
  compareHf(push, 'footer', a.footers, b.footers);
  compareHf(push, 'header', a.headers, b.headers);

  // Field kinds (multiset).
  multisetDiff(push, 'fields', a.fields.map((f) => f.kind), b.fields.map((f) => f.kind), (kind) =>
    ['page-number', 'page-count', 'section-page-count', 'date'].includes(kind) ? 'changed-with-visual-impact' : 'lost');

  // Relationship types (multiset).
  multisetDiff(push, 'relationships', a.relationships.map((r) => r.type), b.relationships.map((r) => r.type), () => 'changed-but-supported');

  // Tables.
  scalar(push, 'tables.count', a.tables.length, b.tables.length, 'changed-with-visual-impact');
  scalar(push, 'tables.gridSpan', sum(a.tables, 'gridSpan'), sum(b.tables, 'gridSpan'), 'changed-with-visual-impact');
  scalar(push, 'tables.vMerge', sum(a.tables, 'vMerge'), sum(b.tables, 'vMerge'), 'changed-with-visual-impact');

  // Drawings / shapes / manual breaks.
  scalar(push, 'drawings.count', a.drawings.length, b.drawings.length, 'changed-with-visual-impact');
  scalar(push, 'shapes.count', a.shapes.length, b.shapes.length, 'lost');
  scalar(push, 'manualPageBreaks', a.manualPageBreaks, b.manualPageBreaks, 'changed-with-visual-impact');

  if (diffs.length === 0) diffs.push({ aspect: 'overall', status: 'equal', from: null, to: null, note: 'No material structural differences detected.' });
  return diffs;
}

function geomDiff(push, i, aspect, from, to) {
  if (from === to) return;
  const deltaTwips = Math.abs(from - to);
  // < 1pt (20 twips) rounding is equivalent; larger is a visible geometry shift.
  const status = deltaTwips <= 20 ? 'equivalent' : 'changed-with-visual-impact';
  push(`section[${i}].${aspect}`, status, from, to, `Δ ${deltaTwips} twips`);
}
function compareHf(push, kind, a, b) {
  const at = a.map((x) => x.text).filter(Boolean);
  const bt = b.map((x) => x.text).filter(Boolean);
  for (const text of at) {
    if (!bt.some((x) => normalize(x) === normalize(text))) {
      // Page-number footers legitimately differ once the field value changes.
      const dynamic = /\bz\b|\d+\s*\/\s*\d+|strona/i.test(text);
      push(`${kind}.text`, dynamic ? 'changed-with-visual-impact' : 'lost', truncate(text, 50), null, dynamic ? 'Footer text differs (may be page-number field value).' : `${kind} text missing in exported document.`);
    }
  }
}
function multisetDiff(push, aspect, a, b, statusFor) {
  const ca = tally(a), cb = tally(b);
  for (const [key, count] of ca) {
    const other = cb.get(key) || 0;
    if (other < count) push(`${aspect}[${key}]`, statusFor(key), count, other, 'Fewer than source (dropped).');
  }
  for (const [key, count] of cb) {
    const other = ca.get(key) || 0;
    if (other < count) push(`${aspect}[${key}]`, 'added', other, count, 'More than source.');
  }
}
function scalar(push, aspect, from, to, lossStatus) {
  if (from === to) return;
  push(aspect, to < from ? lossStatus : 'added', from, to, '');
}

function printConsole(files, diffs, regressions) {
  console.log(`\x1b[1mcompare-docx\x1b[0m\n  source:   ${files[0]}\n  exported: ${files[1]}\n`);
  const color = { equal: 2, equivalent: 2, 'changed-but-supported': 6, 'changed-with-visual-impact': 33, unsupported: 31, lost: 31, added: 35 };
  for (const d of diffs) {
    const c = color[d.status] ?? 0;
    console.log(`  \x1b[${c}m${d.status.padEnd(28)}\x1b[0m ${d.aspect}  ${fmt(d.from)} → ${fmt(d.to)}${d.note ? '  (' + d.note + ')' : ''}`);
  }
  console.log(`\n${regressions.length ? `\x1b[31m${regressions.length} regression(s)\x1b[0m` : '\x1b[32mno material regressions\x1b[0m'}`);
}

function tally(arr) { const m = new Map(); for (const k of arr) m.set(k, (m.get(k) || 0) + 1); return m; }
function sum(arr, key) { return arr.reduce((a, x) => a + (x[key] || 0), 0); }
function normalize(s) { return (s || '').replace(/\s+/g, ' ').trim().toLowerCase(); }
function truncate(s, n) { return (s || '').length > n ? s.slice(0, n) + '…' : s || ''; }
function fmt(v) { return v === null || v === undefined ? '—' : String(v); }

process.exit(main(process.argv));
