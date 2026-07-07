// Minimal, zero-dependency ZIP reader for OOXML (.docx) packages.
//
// A DOCX is a ZIP archive; entries are either stored (method 0) or DEFLATE
// (method 8). We parse the central directory and inflate with Node's built-in
// zlib so the diagnostics tools need no third-party packages (matching this
// repo's zero-dependency conversion stack). The reader never writes to disk.

import { inflateRawSync } from 'node:zlib';

const EOCD_SIG = 0x06054b50; // End Of Central Directory
const CDH_SIG = 0x02014b50; // Central Directory Header
const LFH_SIG = 0x04034b50; // Local File Header

/**
 * @typedef {Object} ZipEntry
 * @property {string} name
 * @property {() => string} text  UTF-8 decoded content (lazy).
 * @property {number} size        uncompressed size in bytes.
 */

/**
 * Read all entries from a ZIP buffer.
 * @param {Buffer} buf
 * @returns {Map<string, ZipEntry>}
 */
export function readZip(buf) {
  const eocd = findEocd(buf);
  if (eocd < 0) throw new Error('Not a ZIP/DOCX package (no EOCD record).');

  const cdCount = buf.readUInt16LE(eocd + 10);
  let ptr = buf.readUInt32LE(eocd + 16);

  const entries = new Map();
  for (let i = 0; i < cdCount; i++) {
    if (buf.readUInt32LE(ptr) !== CDH_SIG) break;
    const method = buf.readUInt16LE(ptr + 10);
    const compSize = buf.readUInt32LE(ptr + 20);
    const uncompSize = buf.readUInt32LE(ptr + 24);
    const nameLen = buf.readUInt16LE(ptr + 28);
    const extraLen = buf.readUInt16LE(ptr + 30);
    const commentLen = buf.readUInt16LE(ptr + 32);
    const localOffset = buf.readUInt32LE(ptr + 42);
    const name = buf.toString('utf8', ptr + 46, ptr + 46 + nameLen);
    ptr += 46 + nameLen + extraLen + commentLen;

    entries.set(name, makeEntry(buf, name, localOffset, method, compSize, uncompSize));
  }
  return entries;
}

/**
 * @param {Buffer} buf
 * @param {string} name
 * @param {number} localOffset
 * @param {number} method
 * @param {number} compSize
 * @param {number} uncompSize
 * @returns {ZipEntry}
 */
function makeEntry(buf, name, localOffset, method, compSize, uncompSize) {
  let cachedText = null;
  return {
    name,
    size: uncompSize,
    text() {
      if (cachedText !== null) return cachedText;
      if (buf.readUInt32LE(localOffset) !== LFH_SIG) {
        throw new Error(`Corrupt local header for "${name}".`);
      }
      const nameLen = buf.readUInt16LE(localOffset + 26);
      const extraLen = buf.readUInt16LE(localOffset + 28);
      const dataStart = localOffset + 30 + nameLen + extraLen;
      const raw = buf.subarray(dataStart, dataStart + compSize);
      const bytes = method === 0 ? raw : inflateRawSync(raw);
      cachedText = bytes.toString('utf8');
      return cachedText;
    },
  };
}

/**
 * @param {Buffer} buf
 * @returns {number} offset of EOCD, or -1.
 */
function findEocd(buf) {
  // EOCD is at the end; scan backwards over the (max 64KB) comment.
  const min = Math.max(0, buf.length - 0xffff - 22);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === EOCD_SIG) return i;
  }
  return -1;
}
