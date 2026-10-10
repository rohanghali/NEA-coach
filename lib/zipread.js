'use strict';
/*
 * zipread.js — a minimal, read-only ZIP reader.
 *
 * WHY THIS EXISTS
 * A .pptx file is really just a ZIP archive full of XML. This module reads the
 * archive directly from disk using only Node's built-in zlib. It opens the file
 * with 'r' (read-only) and never writes back. There is no code path anywhere in
 * NEA Coach that can modify your PowerPoint file.
 *
 * No external dependencies, no native modules, so it behaves identically on
 * Intel and Apple Silicon Macs.
 */

const fs = require('fs');
const zlib = require('zlib');

// ZIP format constants
const SIG_EOCD = 0x06054b50; // End of central directory
const SIG_EOCD64 = 0x06064b50; // ZIP64 end of central directory
const SIG_EOCD64_LOC = 0x07064b50; // ZIP64 EOCD locator
const SIG_CENTRAL = 0x02014b50; // Central directory file header
const SIG_LOCAL = 0x04034b50; // Local file header

const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

// Hard ceiling for any single decompression. A 24-megapixel PNG is about 70 MB
// uncompressed, so 192 MB leaves generous headroom while still refusing to
// allocate a runaway result from a malformed or hostile entry.
const INFLATE_CEILING = 192 * 1024 * 1024;

/**
 * Pick a first-attempt ceiling for a prefix read. Image headers are tiny, but
 * giving zlib a little room means most images succeed on the first attempt.
 */
function prefixCeiling(maxBytes) {
  return Math.min(INFLATE_CEILING, Math.max(64 * 1024, maxBytes * 4));
}

/** Read a little-endian unsigned integer. */
function u16(buf, off) {
  return buf.readUInt16LE(off);
}
function u32(buf, off) {
  return buf.readUInt32LE(off);
}

/**
 * Locate the End Of Central Directory record.
 * It sits at the end of the file, followed by a comment of 0-65535 bytes,
 * so we scan backwards through that window.
 */
function findEocd(buf) {
  const maxComment = 0xffff;
  const minPos = Math.max(0, buf.length - (22 + maxComment));
  for (let i = buf.length - 22; i >= minPos; i--) {
    if (u32(buf, i) === SIG_EOCD) return i;
  }
  return -1;
}

class ZipArchive {
  constructor(filePath) {
    this.filePath = filePath;
    // Read the whole thing. Portfolios of 200 A3 pages with images are still
    // only tens of MB, so this is fast and keeps the logic simple.
    this.buf = fs.readFileSync(filePath);
    this.entries = new Map(); // name -> central directory entry
    this._parse();
  }

  _parse() {
    const buf = this.buf;
    const eocd = findEocd(buf);
    if (eocd < 0) throw new Error('Not a ZIP file (no end-of-central-directory record)');

    let entryCount = u16(buf, eocd + 10);
    let cdOffset = u32(buf, eocd + 16);

    // ZIP64: used when the archive is huge. PowerPoint files rarely hit this,
    // but a 6-year-old iMac with a big portfolio makes it worth supporting.
    if (entryCount === 0xffff || cdOffset === 0xffffffff) {
      const locOff = eocd - 20;
      if (locOff >= 0 && u32(buf, locOff) === SIG_EOCD64_LOC) {
        const z64 = Number(buf.readBigUInt64LE(locOff + 8));
        if (z64 >= 0 && z64 + 56 <= buf.length && u32(buf, z64) === SIG_EOCD64) {
          entryCount = Number(buf.readBigUInt64LE(z64 + 32));
          cdOffset = Number(buf.readBigUInt64LE(z64 + 48));
        }
      }
    }

    let p = cdOffset;
    for (let i = 0; i < entryCount; i++) {
      if (p + 46 > buf.length || u32(buf, p) !== SIG_CENTRAL) break;

      const method = u16(buf, p + 10);
      let compSize = u32(buf, p + 20);
      let uncompSize = u32(buf, p + 24);
      const nameLen = u16(buf, p + 28);
      const extraLen = u16(buf, p + 30);
      const commentLen = u16(buf, p + 32);
      let localOffset = u32(buf, p + 42);

      const name = buf.toString('utf8', p + 46, p + 46 + nameLen);

      // ZIP64 extended information extra field (id 0x0001)
      if (compSize === 0xffffffff || uncompSize === 0xffffffff || localOffset === 0xffffffff) {
        let ep = p + 46 + nameLen;
        const epEnd = ep + extraLen;
        while (ep + 4 <= epEnd) {
          const id = u16(buf, ep);
          const size = u16(buf, ep + 2);
          let fp = ep + 4;
          if (id === 0x0001) {
            if (uncompSize === 0xffffffff) {
              uncompSize = Number(buf.readBigUInt64LE(fp));
              fp += 8;
            }
            if (compSize === 0xffffffff) {
              compSize = Number(buf.readBigUInt64LE(fp));
              fp += 8;
            }
            if (localOffset === 0xffffffff) {
              localOffset = Number(buf.readBigUInt64LE(fp));
              fp += 8;
            }
          }
          ep += 4 + size;
        }
      }

      this.entries.set(name, { name, method, compSize, uncompSize, localOffset });
      p += 46 + nameLen + extraLen + commentLen;
    }
  }

  /** List every path stored in the archive. */
  list() {
    return Array.from(this.entries.keys());
  }

  has(name) {
    return this.entries.has(name);
  }

  /**
   * Extract one entry's bytes. Returns null if the entry is missing or the
   * data is corrupt, so callers can degrade gracefully instead of crashing.
   *
   * @param {string} name
   * @param {number} [maxBytes] Only decompress about this many bytes.
   *
   * On a real portfolio this matters. Rohan's GCSE deck is 528 MB, almost all of
   * it 344 images, and we only need the first few dozen bytes of each one to read
   * its pixel dimensions. Node's inflateSync cannot stop early: maxOutputLength
   * throws ERR_BUFFER_TOO_LARGE rather than truncating. So when a caller wants a
   * prefix we decompress with a small ceiling first and only fall back to a full
   * inflate if the entry genuinely needs more. Most images finish inside the cap,
   * and we never allocate the whole media set.
   */
  read(name, maxBytes) {
    const e = this.entries.get(name);
    if (!e) return null;

    const buf = this.buf;
    const lo = e.localOffset;
    if (lo + 30 > buf.length || u32(buf, lo) !== SIG_LOCAL) return null;

    const nameLen = u16(buf, lo + 26);
    const extraLen = u16(buf, lo + 28);
    const dataStart = lo + 30 + nameLen + extraLen;
    const dataEnd = dataStart + e.compSize;
    if (dataEnd > buf.length) return null;

    const raw = buf.subarray(dataStart, dataEnd);

    if (e.method === METHOD_STORE) {
      return Buffer.from(maxBytes ? raw.subarray(0, maxBytes) : raw);
    }
    if (e.method !== METHOD_DEFLATE) return null; // unsupported compression

    if (!maxBytes) {
      try {
        return zlib.inflateRawSync(raw, { maxOutputLength: INFLATE_CEILING });
      } catch (err) {
        return null;
      }
    }

    // Two-tier partial read: try a generous ceiling for the header we need, then
    // fall back to a full inflate only for entries that exceed it.
    const props = [prefixCeiling(maxBytes), INFLATE_CEILING];
    for (let i = 0; i < props.length; i++) {
      try {
        const out = zlib.inflateRawSync(raw, { maxOutputLength: props[i] });
        return out.subarray(0, maxBytes);
      } catch (err) {
        // ERR_BUFFER_TOO_LARGE simply means "try the next, larger ceiling".
        if (err && err.code !== 'ERR_BUFFER_TOO_LARGE') return null;
      }
    }
    return null;
  }

  /** Extract an entry as UTF-8 text, or null. */
  readText(name) {
    const b = this.read(name);
    return b ? b.toString('utf8') : null;
  }
}

/**
 * Open a .pptx. We touch the file read-only and hold it in memory; the file on
 * disk is never opened for writing.
 */
function open(filePath) {
  return new ZipArchive(filePath);
}

module.exports = { open, ZipArchive };
