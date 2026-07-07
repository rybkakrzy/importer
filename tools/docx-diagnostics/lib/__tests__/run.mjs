// Self-test for the DOCX diagnostics — zero deps, run with `node run.mjs`.
// Exercises the ZIP reader (inflate path), the report builder, and the compare
// classifier against synthesized fixtures.

import assert from 'node:assert/strict';
import { readZip } from '../zip.mjs';
import { buildReport } from '../docx-report.mjs';
import { sampleSource, sampleExportedWithLoss } from './fixtures.mjs';

let passed = 0;
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  \x1b[32m✓\x1b[0m ${name}`);
  } catch (e) {
    console.error(`  \x1b[31m✗ ${name}\x1b[0m\n    ${e.message}`);
    process.exitCode = 1;
  }
}

console.log('docx-diagnostics self-test');

const report = buildReport(readZip(sampleSource()));

test('reads two sections with correct A4 portrait + landscape geometry', () => {
  assert.equal(report.sections.length, 2);
  assert.equal(report.pageSettings[0].orientation, 'portrait');
  assert.equal(report.pageSettings[0].physicalPageWidth.twips, 11906);
  assert.ok(Math.abs(report.pageSettings[0].physicalPageWidth.cm - 21) < 0.01, 'A4 width ≈ 21 cm');
  assert.equal(report.pageSettings[1].orientation, 'landscape');
  assert.equal(report.pageSettings[1].physicalPageWidth.twips, 16838);
});

test('computes content area = page minus margins', () => {
  const ps = report.pageSettings[0];
  assert.equal(ps.contentAreaWidth.twips, 11906 - 1417 - 1417);
});

test('extracts the footer part text and PAGE/NUMPAGES fields', () => {
  const footer = report.footers.find((f) => f.file === 'word/footer1.xml');
  assert.ok(footer, 'footer1.xml analyzed');
  assert.match(footer.text, /Strona:/);
  assert.match(footer.text, /Przedluzenie/);
  const kinds = report.fields.map((f) => f.kind);
  assert.ok(kinds.includes('page-number'), 'PAGE detected');
  assert.ok(kinds.includes('page-count'), 'NUMPAGES detected');
});

test('resolves footer references to a part (not missing)', () => {
  assert.ok(report.footerReferences.length >= 1);
  assert.equal(report.footerReferences.every((r) => !r.missing), true);
});

test('reports the table with gridSpan and vMerge', () => {
  assert.equal(report.tables.length, 1);
  assert.equal(report.tables[0].gridSpan, 1);
  assert.equal(report.tables[0].vMerge, 1);
});

test('flags a footnote reference as a data-loss-risk element', () => {
  const fn = report.unsupportedElements.find((u) => u.element === 'w:footnoteReference');
  assert.ok(fn, 'footnote reference flagged');
  assert.equal(fn.risk, 'data-loss');
});

test('counts the manual page break', () => {
  assert.equal(report.manualPageBreaks, 1);
});

// --- Broken export: missing footer part + lost section --------------------
const broken = buildReport(readZip(sampleExportedWithLoss()));

test('detects a dangling footer reference as a critical warning', () => {
  assert.ok(broken.warnings.some((w) => w.severity === 'critical'), 'critical warning present');
  assert.ok(broken.footerReferences.some((r) => r.missing), 'missing footer reference');
});

console.log(`\n${passed} passed${process.exitCode ? ', with failures' : ''}`);
