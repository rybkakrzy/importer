// Builds a structured diagnostic report from a DOCX package. Regex-based OOXML
// extraction (zero dependencies) — sufficient for diagnostics; it reports the
// structure the D2 importer cares about and flags elements the importer is
// known to drop, so nobody loses data unknowingly.

const TWIPS_PER_CM = 566.9291338582677; // 1 cm = 567 twips (1440 twip/inch / 2.54)
const TWIPS_PER_PT = 20;

/** Elements the D2 reader silently drops (from .ai/AUDIT_WORD_COMPATIBILITY.md). */
const UNSUPPORTED_PROBES = [
  ['w:footnoteReference', 'footnotes (KR-02) — content lost on autosave'],
  ['w:endnoteReference', 'endnotes (KR-02) — content lost on autosave'],
  ['w:commentReference', 'comments (KR-03) — lost on autosave'],
  ['w:ins', 'tracked insertions (KR-04) — inserted text not rendered'],
  ['w:del', 'tracked deletions (KR-04) — silently accepted'],
  ['m:oMath', 'equations (KR-06) — dropped'],
  ['wps:txbx', 'modern text boxes (KR-06) — only text extracted'],
  ['v:textbox', 'legacy VML text boxes (KR-06)'],
  ['w:bookmarkStart', 'bookmarks (KR-07) — anchors lost'],
  ['w:cols', 'multi-column layout (KR-09) — rendered single column'],
  ['w:pgNumType', 'section page numbering (KR-10) — lost'],
  ['w:vanish', 'hidden text (KR-12) — becomes visible / attribute lost'],
  ['a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart', 'charts — dropped'],
  ['w:documentProtection', 'document protection (KR-11) — lost'],
];

/**
 * @param {Map<string, {name:string,text:()=>string,size:number}>} entries
 */
export function buildReport(entries) {
  const report = {
    sections: [],
    pageSettings: [],
    headers: [],
    footers: [],
    footerReferences: [],
    fields: [],
    tables: [],
    paragraphSpacing: [],
    drawings: [],
    shapes: [],
    relationships: [],
    unsupportedElements: [],
    warnings: [],
    manualPageBreaks: 0,
  };

  const docEntry = entries.get('word/document.xml');
  if (!docEntry) {
    report.warnings.push({ severity: 'critical', path: 'word/document.xml', message: 'Main document part missing — not a valid WordprocessingML package.' });
    return report;
  }

  let docXml;
  try {
    docXml = docEntry.text();
  } catch (e) {
    report.warnings.push({ severity: 'critical', path: 'word/document.xml', message: `Unreadable: ${e.message}` });
    return report;
  }

  const rels = parseRels(entries, 'word/_rels/document.xml.rels');
  report.relationships = [...rels.values()].map((r) => ({ id: r.id, type: shortRelType(r.type), target: r.target, external: r.external }));

  // --- Sections + page settings ------------------------------------------
  const sectPrs = matchAll(docXml, /<w:sectPr\b[^>]*(?:\/>|[\s\S]*?<\/w:sectPr>)/g);
  sectPrs.forEach((sectXml, idx) => {
    const pgSz = firstMatch(sectXml, /<w:pgSz\b[^>]*\/?>/);
    const pgMar = firstMatch(sectXml, /<w:pgMar\b[^>]*\/?>/);
    const w = num(attr(pgSz, 'w:w'));
    const h = num(attr(pgSz, 'w:h'));
    const orient = attr(pgSz, 'w:orient') || (w && h && w > h ? 'landscape' : 'portrait');
    const cols = num(attr(firstMatch(sectXml, /<w:cols\b[^>]*\/?>/), 'w:num')) || 1;

    report.sections.push({
      sectionIndex: idx,
      orientation: orient,
      columns: cols,
      titlePage: /<w:titlePg\b/.test(sectXml),
      headerReferences: parseHfRefs(sectXml, 'header', rels),
      footerReferences: parseHfRefs(sectXml, 'footer', rels),
    });

    report.pageSettings.push({
      sectionIndex: idx,
      physicalPageWidth: dims(w),
      physicalPageHeight: dims(h),
      orientation: orient,
      margins: {
        top: dims(num(attr(pgMar, 'w:top'))),
        right: dims(num(attr(pgMar, 'w:right'))),
        bottom: dims(num(attr(pgMar, 'w:bottom'))),
        left: dims(num(attr(pgMar, 'w:left'))),
        header: dims(num(attr(pgMar, 'w:header'))),
        footer: dims(num(attr(pgMar, 'w:footer'))),
        gutter: dims(num(attr(pgMar, 'w:gutter'))),
      },
      contentAreaWidth: dims(sub(w, num(attr(pgMar, 'w:left')), num(attr(pgMar, 'w:right')))),
      contentAreaHeight: dims(sub(h, num(attr(pgMar, 'w:top')), num(attr(pgMar, 'w:bottom')))),
    });

    report.footerReferences.push(...parseHfRefs(sectXml, 'footer', rels).map((r) => ({ ...r, sectionIndex: idx })));
  });

  // --- Header/footer parts ------------------------------------------------
  for (const name of [...entries.keys()].sort()) {
    const m = /^word\/(header|footer)\d*\.xml$/.exec(name);
    if (!m) continue;
    const analyzed = analyzePart(entries, name);
    (m[1] === 'header' ? report.headers : report.footers).push(analyzed);
  }

  // --- Fields, tables, spacing, drawings, shapes across doc + h/f ---------
  const partsToScan = ['word/document.xml', ...report.headers.map((h) => h.file), ...report.footers.map((f) => f.file)];
  for (const partName of partsToScan) {
    const entry = entries.get(partName);
    if (!entry) continue;
    let xml;
    try { xml = entry.text(); } catch { continue; }
    report.fields.push(...extractFields(xml).map((f) => ({ ...f, part: partName })));
  }

  report.manualPageBreaks = countOccurrences(docXml, 'w:type="page"');
  report.tables = extractTables(docXml).map((t) => ({ ...t, part: 'word/document.xml' }));
  report.paragraphSpacing = extractSpacing(docXml);
  report.drawings = extractDrawings(docXml, 'word/document.xml');
  report.shapes = extractShapes(docXml, 'word/document.xml');

  // --- Unsupported elements (data-loss risk) ------------------------------
  for (const partName of partsToScan) {
    const entry = entries.get(partName);
    if (!entry) continue;
    let xml;
    try { xml = entry.text(); } catch { continue; }
    for (const [probe, note] of UNSUPPORTED_PROBES) {
      const count = countElementLike(xml, probe);
      if (count > 0) {
        report.unsupportedElements.push({ element: probe.split(' ')[0], count, part: partName, note, risk: 'data-loss' });
      }
    }
  }

  // --- Relationship integrity warnings ------------------------------------
  for (const ref of report.footerReferences) {
    if (ref.missing) {
      report.warnings.push({ severity: 'critical', path: `word/_rels/document.xml.rels#${ref.relationshipId}`, message: `Footer reference "${ref.relationshipId}" (${ref.type}) has no target part.` });
    }
  }
  for (const section of report.sections) {
    for (const ref of section.headerReferences) {
      if (ref.missing) report.warnings.push({ severity: 'critical', path: `section ${section.sectionIndex}`, message: `Header reference "${ref.relationshipId}" (${ref.type}) has no target part.` });
    }
  }

  return report;
}

// ---------------------------------------------------------------------------
// Extraction helpers
// ---------------------------------------------------------------------------

function analyzePart(entries, name) {
  let xml = '';
  try { xml = entries.get(name).text(); } catch { /* corrupt */ }
  const relPath = name.replace(/(^|\/)([^/]+)$/, '$1_rels/$2.rels');
  const rels = parseRels(entries, relPath);
  const imageRels = [...rels.values()].filter((r) => /image|media/i.test(r.type));
  return {
    file: name,
    type: /header/.test(name) ? 'header' : 'footer',
    text: textOf(xml),
    fields: extractFields(xml).map((f) => f.kind),
    tableCount: countOccurrences(xml, '<w:tbl>'),
    imageCount: countOccurrences(xml, '<a:blip') + countOccurrences(xml, '<v:imagedata'),
    drawingCount: countOccurrences(xml, '<w:drawing'),
    shapeCount: countOccurrences(xml, '<wps:wsp') + countOccurrences(xml, '<v:shape'),
    relationships: imageRels.map((r) => ({ id: r.id, target: r.target, missing: !entries.has(joinRel(name, r.target)) && !r.external })),
  };
}

function parseRels(entries, relPath) {
  const map = new Map();
  const entry = entries.get(relPath);
  if (!entry) return map;
  let xml = '';
  try { xml = entry.text(); } catch { return map; }
  for (const rel of matchAll(xml, /<Relationship\b[^>]*\/?>/g)) {
    const id = attr(rel, 'Id');
    if (!id) continue;
    map.set(id, { id, type: attr(rel, 'Type') || '', target: attr(rel, 'Target') || '', external: /External/i.test(attr(rel, 'TargetMode') || '') });
  }
  return map;
}

function parseHfRefs(sectXml, kind, rels) {
  const tag = kind === 'header' ? 'w:headerReference' : 'w:footerReference';
  return matchAll(sectXml, new RegExp(`<${tag}\\b[^>]*\\/?>`, 'g')).map((ref) => {
    const relId = attr(ref, 'r:id');
    const type = attr(ref, 'w:type') || 'default';
    const rel = rels.get(relId);
    return {
      relationshipId: relId,
      type,
      target: rel ? rel.target : null,
      xmlFile: rel ? joinRel('word/document.xml', rel.target) : null,
      missing: !rel,
    };
  });
}

function extractFields(xml) {
  const fields = [];
  // Simple fields.
  for (const f of matchAll(xml, /<w:fldSimple\b[^>]*>/g)) {
    fields.push({ kind: classifyInstr(attr(f, 'w:instr')), instruction: (attr(f, 'w:instr') || '').trim(), form: 'simple' });
  }
  // Complex fields: concatenate instrText between begin/end.
  const instrTexts = matchAll(xml, /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g).map((m) => stripTags(m));
  for (const instr of instrTexts) {
    fields.push({ kind: classifyInstr(instr), instruction: instr.trim(), form: 'complex' });
  }
  return fields;
}

function classifyInstr(instr) {
  const s = (instr || '').trim().toUpperCase();
  if (/^PAGE\b/.test(s)) return 'page-number';
  if (/^NUMPAGES\b/.test(s)) return 'page-count';
  if (/^SECTIONPAGES\b/.test(s)) return 'section-page-count';
  if (/^SECTION\b/.test(s)) return 'section-number';
  if (/^(DATE|TIME|CREATEDATE|SAVEDATE)\b/.test(s)) return 'date';
  if (/^TOC\b/.test(s)) return 'toc';
  if (/^(REF|PAGEREF)\b/.test(s)) return 'ref';
  if (/^HYPERLINK\b/.test(s)) return 'hyperlink';
  if (/^MERGEFIELD\b/.test(s)) return 'mergefield';
  if (s.length === 0) return 'unknown';
  return 'unsupported';
}

function extractTables(xml) {
  return matchAll(xml, /<w:tbl>[\s\S]*?<\/w:tbl>/g).map((tbl, i) => ({
    index: i,
    rows: countOccurrences(tbl, '<w:tr>') + countOccurrences(tbl, '<w:tr '),
    gridColumns: countOccurrences(tbl, '<w:gridCol'),
    gridSpan: countOccurrences(tbl, '<w:gridSpan'),
    vMerge: countOccurrences(tbl, '<w:vMerge'),
    nested: /<w:tbl>[\s\S]*<w:tbl>/.test(tbl),
    styled: /<w:tblStyle\b/.test(tbl),
  }));
}

function extractSpacing(xml) {
  const out = [];
  const seen = new Set();
  for (const sp of matchAll(xml, /<w:spacing\b[^>]*\/?>/g)) {
    const key = sp;
    if (seen.has(key)) continue;
    seen.add(key);
    const line = attr(sp, 'w:line');
    const rule = attr(sp, 'w:lineRule');
    out.push({ before: attr(sp, 'w:before'), after: attr(sp, 'w:after'), line, lineRule: rule || (line ? 'auto' : null) });
    if (out.length >= 50) break;
  }
  return out;
}

function extractDrawings(xml, part) {
  return matchAll(xml, /<w:drawing>[\s\S]*?<\/w:drawing>/g).map((d, i) => ({
    index: i,
    part,
    mode: /<wp:anchor\b/.test(d) ? 'anchored' : 'inline',
    hasImage: /<a:blip\b/.test(d),
    hasChart: /uri="[^"]*chart/.test(d),
    hasTextbox: /<wps:txbx\b/.test(d),
    wrap: (firstMatch(d, /<wp:wrap\w+\b/) || '').replace(/[<]/, '') || 'none',
  }));
}

function extractShapes(xml, part) {
  const shapes = [];
  for (const s of matchAll(xml, /<wps:wsp>[\s\S]*?<\/wps:wsp>/g)) {
    shapes.push({ part, format: 'drawingml', hasTextbox: /<wps:txbx\b/.test(s), hasTable: /<w:tbl>/.test(s) });
  }
  for (const s of matchAll(xml, /<v:shape\b[\s\S]*?<\/v:shape>/g)) {
    shapes.push({ part, format: 'vml', hasTextbox: /<v:textbox\b/.test(s), hasTable: /<w:tbl>/.test(s) });
  }
  return shapes;
}

// ---------------------------------------------------------------------------
// Small XML/text utilities
// ---------------------------------------------------------------------------

function textOf(xml) {
  return matchAll(xml, /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g).map(stripTags).join('').replace(/\s+/g, ' ').trim();
}
function stripTags(fragment) {
  const inner = fragment.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '');
  return inner.replace(/<[^>]*>/g, '');
}
function matchAll(xml, re) {
  const out = [];
  let m;
  const r = new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g');
  while ((m = r.exec(xml)) !== null) { out.push(m[1] !== undefined ? m[0] : m[0]); if (m.index === r.lastIndex) r.lastIndex++; }
  return out;
}
function firstMatch(xml, re) { const m = (xml || '').match(re); return m ? m[0] : ''; }
function attr(fragment, name) {
  if (!fragment) return '';
  const m = fragment.match(new RegExp(`${name.replace(':', '\\:')}="([^"]*)"`));
  return m ? m[1] : '';
}
function countOccurrences(xml, needle) {
  if (!xml) return 0;
  let count = 0, idx = 0;
  while ((idx = xml.indexOf(needle, idx)) !== -1) { count++; idx += needle.length; }
  return count;
}
/**
 * Count element occurrences by opening tag, respecting name boundaries so
 * "w:ins" does not match "w:instrText". Attribute-style probes (containing a
 * space or '=') fall back to a plain substring count.
 */
function countElementLike(xml, probe) {
  if (!xml) return 0;
  if (/[ =]/.test(probe)) return countOccurrences(xml, probe);
  const opener = '<' + probe;
  let count = 0, idx = 0;
  while ((idx = xml.indexOf(opener, idx)) !== -1) {
    const next = xml.charAt(idx + opener.length);
    if (!/[A-Za-z0-9]/.test(next)) count++; // boundary: space, '>', '/', etc.
    idx += opener.length;
  }
  return count;
}
function num(v) { const n = parseInt(v, 10); return Number.isFinite(n) ? n : 0; }
function sub(total, ...parts) { return total ? total - parts.reduce((a, b) => a + (b || 0), 0) : 0; }
function dims(twips) {
  return { twips: twips || 0, cm: round((twips || 0) / TWIPS_PER_CM, 3), pt: round((twips || 0) / TWIPS_PER_PT, 1) };
}
function round(n, d) { const f = 10 ** d; return Math.round(n * f) / f; }
function shortRelType(type) { const m = /\/([^/]+)$/.exec(type || ''); return m ? m[1] : type; }
function joinRel(basePart, target) {
  if (/^\//.test(target)) return target.slice(1);
  const baseDir = basePart.replace(/\/[^/]*$/, '');
  const stack = baseDir.split('/');
  for (const seg of target.split('/')) {
    if (seg === '..') stack.pop();
    else if (seg !== '.') stack.push(seg);
  }
  return stack.join('/');
}
