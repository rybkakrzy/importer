// Zero-dependency DOCX fixture builder for the diagnostics self-test.
// Includes a minimal ZIP writer (DEFLATE, method 8) so the reader's inflate
// path is exercised too. Fixtures are synthesized (no real customer files).

import { deflateRawSync } from 'node:zlib';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * @param {Record<string,string>} files  name → UTF-8 content
 * @returns {Buffer} a valid ZIP/DOCX buffer
 */
export function makeZip(files) {
  const locals = [];
  const central = [];
  let offset = 0;

  for (const [name, content] of Object.entries(files)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const data = Buffer.from(content, 'utf8');
    const comp = deflateRawSync(data);
    const crc = crc32(data);

    const lfh = Buffer.alloc(30);
    lfh.writeUInt32LE(0x04034b50, 0);
    lfh.writeUInt16LE(20, 4); // version needed
    lfh.writeUInt16LE(0, 6); // flags
    lfh.writeUInt16LE(8, 8); // method: deflate
    lfh.writeUInt16LE(0, 10); // time
    lfh.writeUInt16LE(0, 12); // date
    lfh.writeUInt32LE(crc, 14);
    lfh.writeUInt32LE(comp.length, 18);
    lfh.writeUInt32LE(data.length, 22);
    lfh.writeUInt16LE(nameBuf.length, 26);
    lfh.writeUInt16LE(0, 28); // extra len
    locals.push(lfh, nameBuf, comp);

    const cdh = Buffer.alloc(46);
    cdh.writeUInt32LE(0x02014b50, 0);
    cdh.writeUInt16LE(20, 4);
    cdh.writeUInt16LE(20, 6);
    cdh.writeUInt16LE(0, 8);
    cdh.writeUInt16LE(8, 10);
    cdh.writeUInt16LE(0, 12);
    cdh.writeUInt16LE(0, 14);
    cdh.writeUInt32LE(crc, 16);
    cdh.writeUInt32LE(comp.length, 20);
    cdh.writeUInt32LE(data.length, 24);
    cdh.writeUInt16LE(nameBuf.length, 28);
    cdh.writeUInt32LE(offset, 42);
    central.push(cdh, nameBuf);

    offset += lfh.length + nameBuf.length + comp.length;
  }

  const centralBuf = Buffer.concat(central);
  const localBuf = Buffer.concat(locals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(Object.keys(files).length, 8);
  eocd.writeUInt16LE(Object.keys(files).length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(localBuf.length, 16);
  return Buffer.concat([localBuf, centralBuf, eocd]);
}

const RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
</Relationships>`;

// Footer matching the prompt's report: "Strona: 1 z 1. Przedłużenie..." built
// from PAGE + NUMPAGES complex fields.
const FOOTER1 = `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t xml:space="preserve">Strona: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> z </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t>. Przedluzenie tekstu stopki</w:t></w:r>
  </w:p>
</w:ftr>`;

// Body: a table (gridSpan + vMerge), a manual page break, a footnote ref
// (unsupported → data-loss), and two sections (portrait + landscape).
const DOCUMENT = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Section one paragraph.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:t>Spaced paragraph.</w:t></w:r></w:p>
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:footerReference w:type="default" r:id="rId1"/>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
        </w:sectPr>
      </w:pPr>
    </w:p>
    <w:p><w:r><w:t>Landscape annexe.</w:t></w:r></w:p>
    <w:sectPr>
      <w:footerReference w:type="default" r:id="rId1"/>
      <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

/** A representative source DOCX. */
export function sampleSource() {
  return makeZip({
    'word/document.xml': DOCUMENT,
    'word/_rels/document.xml.rels': RELS,
    'word/footer1.xml': FOOTER1,
  });
}

// A single-section "exported" document that still references footer rId1 but
// whose footer part and relationship were lost — reproduces the regression the
// tool must catch (flattened to one section + dangling footer reference).
const DOCUMENT_LOST = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Section one paragraph.</w:t></w:r></w:p>
    <w:sectPr>
      <w:footerReference w:type="default" r:id="rId1"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

/** An "exported" variant that lost the footer part and a section (regression). */
export function sampleExportedWithLoss() {
  return makeZip({
    'word/document.xml': DOCUMENT_LOST,
    // rId1 (footer) relationship removed; footer1.xml part omitted → dangling ref.
    'word/_rels/document.xml.rels': RELS.replace(/<Relationship Id="rId1"[^>]*\/>/, ''),
  });
}
