'use strict';
/*
 * pdftext.js — a minimal, zero-dependency PDF text extractor.
 *
 * WHY THIS EXISTS
 * The app runs its free structural checks by reading text out of a .pptx. When
 * a student only has a PDF (a Google Slides or PowerPoint export) those checks
 * could not run at all. This module pulls the text out of each page of a PDF so
 * the same checks can run on an exported deck.
 *
 * Everything here uses only Node built-ins: fs, zlib, Buffer. There are no
 * third-party packages, no native modules, no shelling out and no macOS- or
 * Windows-specific tools, so it behaves identically on both platforms.
 *
 * It is deliberately forgiving. A malformed PDF does not throw: whatever could
 * be read is returned, with a short human-readable note in `warnings`.
 */

const fs = require('fs');
const zlib = require('zlib');

// Refuse to allocate an absurd decompression result from a hostile stream.
const INFLATE_CEILING = 256 * 1024 * 1024;

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function isWS(c) {
  return c === ' ' || c === '\n' || c === '\r' || c === '\t' || c === '\f' || c === '\0';
}
function isDelim(c) {
  return (
    c === '(' || c === ')' || c === '<' || c === '>' || c === '[' || c === ']' ||
    c === '{' || c === '}' || c === '/' || c === '%'
  );
}
function isRegular(c) {
  return c !== undefined && !isWS(c) && !isDelim(c);
}

// Value predicates. Names, strings and refs are tagged objects so they can be
// told apart from arrays (also objects) and dicts (plain objects).
function isName(v) { return v !== null && typeof v === 'object' && typeof v.name === 'string'; }
function isRef(v) { return v !== null && typeof v === 'object' && typeof v.ref === 'number'; }
function isPdfString(v) { return v !== null && typeof v === 'object' && typeof v.str === 'string'; }
function isStream(v) { return v !== null && typeof v === 'object' && !!v.stream; }

function nameValue(v) { return isName(v) ? v.name : undefined; }

/** Decode a hex string body ("48656C6C6F") into a latin1 byte string. */
function hexToBytes(hex) {
  const clean = hex.replace(/[^0-9A-Fa-f]/g, '');
  const pairs = Math.floor(clean.length / 2);
  let out = '';
  for (let i = 0; i < pairs; i++) {
    out += String.fromCharCode(parseInt(clean.substr(i * 2, 2), 16));
  }
  if (clean.length % 2 === 1) {
    // A trailing odd nibble means a final byte padded with 0.
    out += String.fromCharCode(parseInt(clean[clean.length - 1] + '0', 16));
  }
  return out;
}

/** Decode a hex string containing UTF-16BE code units into a JS string. */
function hexToUtf16(hex) {
  const clean = hex.replace(/[^0-9A-Fa-f]/g, '');
  let out = '';
  for (let i = 0; i + 3 < clean.length + 1 && i + 2 <= clean.length; i += 4) {
    const chunk = clean.substr(i, 4);
    if (chunk.length < 4) break;
    out += String.fromCharCode(parseInt(chunk, 16));
  }
  return out;
}

// ---------------------------------------------------------------------------
// Object syntax parser
//
// Everything is parsed out of a latin1 view of the file so that string offsets
// are byte offsets, which makes slicing raw stream data exact.
// ---------------------------------------------------------------------------

function skipWS(s, p) {
  while (p < s.length) {
    const c = s[p];
    if (isWS(c)) { p++; continue; }
    if (c === '%') {
      while (p < s.length && s[p] !== '\n' && s[p] !== '\r') p++;
      continue;
    }
    break;
  }
  return p;
}

function readRegular(s, p) {
  let out = '';
  while (p < s.length && isRegular(s[p])) { out += s[p]; p++; }
  return { token: out, p };
}

/** Parse one object; returns { v, p } or null when nothing sensible is there. */
function parseValue(s, p) {
  p = skipWS(s, p);
  if (p >= s.length) return null;
  const c = s[p];

  if (c === '<') {
    if (s[p + 1] === '<') return parseDict(s, p);
    const end = s.indexOf('>', p + 1);
    if (end < 0) return { v: { str: '' }, p: s.length };
    return { v: { str: hexToBytes(s.slice(p + 1, end)) }, p: end + 1 };
  }

  if (c === '(') return parseLiteralString(s, p);

  if (c === '[') {
    p++;
    const arr = [];
    for (;;) {
      p = skipWS(s, p);
      if (p >= s.length) break;
      if (s[p] === ']') { p++; break; }
      const r = parseValue(s, p);
      if (!r) break;
      arr.push(r.v);
      if (r.p === p) { p++; continue; } // safety: never spin
      p = r.p;
    }
    return { v: arr, p };
  }

  if (c === '/') {
    p++;
    const r = readRegular(s, p);
    return { v: { name: r.token }, p: r.p };
  }

  // Numbers, booleans, null, and indirect references.
  if (c === '+' || c === '-' || c === '.' || (c >= '0' && c <= '9')) {
    const r = readRegular(s, p);
    const num = Number(r.token);
    if (!isFinite(num)) return { v: 0, p: r.p };
    // Look ahead for "gen R" to spot an indirect reference.
    const after = skipWS(s, r.p);
    const r2 = readRegular(s, after);
    if (/^\d+$/.test(r2.token) && Number.isInteger(num) && num >= 0) {
      const after2 = skipWS(s, r2.p);
      if (s[after2] === 'R' && !isRegular(s[after2 + 1])) {
        return { v: { ref: num, gen: Number(r2.token) }, p: after2 + 1 };
      }
    }
    return { v: num, p: r.p };
  }

  const r = readRegular(s, p);
  if (r.token === 'true') return { v: true, p: r.p };
  if (r.token === 'false') return { v: false, p: r.p };
  if (r.token === 'null') return { v: null, p: r.p };
  if (r.token === '') return null;
  return { v: { keyword: r.token }, p: r.p };
}

function parseDict(s, p) {
  p += 2; // consume <<
  const dict = {};
  for (;;) {
    p = skipWS(s, p);
    if (p >= s.length) break;
    if (s[p] === '>' && s[p + 1] === '>') { p += 2; break; }
    if (s[p] !== '/') {
      // Not a name; skip one value to try to resynchronise.
      const r = parseValue(s, p);
      if (!r || r.p === p) { p++; continue; }
      p = r.p;
      continue;
    }
    const kr = parseValue(s, p);
    if (!kr) break;
    p = skipWS(s, kr.p);
    const vr = parseValue(s, p);
    if (!vr) break;
    dict[kr.v.name] = vr.v;
    if (vr.p === p) { p++; continue; }
    p = vr.p;
  }
  return { v: dict, p };
}

/** Parse a PDF literal string ( ... ), handling escapes and nesting. */
function parseLiteralString(s, p) {
  p++; // consume (
  let depth = 1;
  let out = '';
  while (p < s.length) {
    const c = s[p];
    if (c === '\\') {
      const n = s[p + 1];
      p += 2;
      switch (n) {
        case 'n': out += '\n'; break;
        case 'r': out += '\r'; break;
        case 't': out += '\t'; break;
        case 'b': out += '\b'; break;
        case 'f': out += '\f'; break;
        case '(': out += '('; break;
        case ')': out += ')'; break;
        case '\\': out += '\\'; break;
        case '\r':
          if (s[p] === '\n') p++; // line continuation
          break;
        case '\n':
          break;
        default:
          if (n >= '0' && n <= '7') {
            let oct = n;
            while (oct.length < 3 && s[p] >= '0' && s[p] <= '7') { oct += s[p]; p++; }
            out += String.fromCharCode(parseInt(oct, 8) & 0xff);
          } else if (n !== undefined) {
            out += n;
          }
      }
      continue;
    }
    if (c === '(') { depth++; out += c; p++; continue; }
    if (c === ')') {
      depth--;
      p++;
      if (depth === 0) break;
      out += c;
      continue;
    }
    out += c;
    p++;
  }
  return { v: { str: out }, p };
}

// ---------------------------------------------------------------------------
// Stream decoding (filters + predictors)
// ---------------------------------------------------------------------------

function pngUnfilter(data, colors, bpc, columns) {
  const bpp = Math.max(1, Math.ceil((colors * bpc) / 8));
  const rowLen = Math.ceil((colors * bpc * columns) / 8);
  const stride = rowLen + 1;
  const rows = Math.floor(data.length / stride);
  const out = Buffer.alloc(rows * rowLen);
  let prev = Buffer.alloc(rowLen);
  for (let r = 0; r < rows; r++) {
    const ft = data[r * stride];
    const src = data.subarray(r * stride + 1, r * stride + 1 + rowLen);
    const cur = Buffer.alloc(rowLen);
    for (let i = 0; i < rowLen; i++) {
      const raw = src[i];
      const a = i >= bpp ? cur[i - bpp] : 0;
      const b = prev[i];
      const cc = i >= bpp ? prev[i - bpp] : 0;
      let val;
      switch (ft) {
        case 1: val = raw + a; break;
        case 2: val = raw + b; break;
        case 3: val = raw + ((a + b) >> 1); break;
        case 4: {
          const pp = a + b - cc;
          const pa = Math.abs(pp - a);
          const pb = Math.abs(pp - b);
          const pc = Math.abs(pp - cc);
          val = raw + (pa <= pb && pa <= pc ? a : pb <= pc ? b : cc);
          break;
        }
        default: val = raw;
      }
      cur[i] = val & 0xff;
    }
    cur.copy(out, r * rowLen);
    prev = cur;
  }
  return out;
}

function tiffUnfilter(data, colors, bpc, columns) {
  const rowLen = Math.ceil((colors * bpc * columns) / 8);
  const out = Buffer.from(data);
  if (bpc !== 8) return out; // only the common 8-bit case is handled
  const rows = Math.floor(out.length / rowLen);
  for (let r = 0; r < rows; r++) {
    const base = r * rowLen;
    for (let i = colors; i < rowLen; i++) {
      out[base + i] = (out[base + i] + out[base + i - colors]) & 0xff;
    }
  }
  return out;
}

function applyPredictor(buf, parms, pdf) {
  const predictor = pdf ? numberOr(pdf.resolve(parms.Predictor), 1) : numberOr(parms.Predictor, 1);
  if (!predictor || predictor === 1) return buf;
  const colors = numberOr(pdf && pdf.resolve(parms.Colors), 1);
  const bpc = numberOr(pdf && pdf.resolve(parms.BitsPerComponent), 8);
  const columns = numberOr(pdf && pdf.resolve(parms.Columns), 1);
  try {
    if (predictor === 2) return tiffUnfilter(buf, colors, bpc, columns);
    if (predictor >= 10) return pngUnfilter(buf, colors, bpc, columns);
  } catch (e) {
    return buf;
  }
  return buf;
}

function numberOr(v, d) {
  return typeof v === 'number' && isFinite(v) ? v : d;
}

/**
 * Decode a stream object's raw bytes. Returns a latin1 string of the decoded
 * bytes; on any failure the raw bytes are returned unchanged so callers still
 * have something to work with.
 */
function decodeStream(pdf, streamObj) {
  let data = Buffer.from(streamObj.data, 'latin1');
  let filters = pdf.resolve(streamObj.stream.Filter);
  if (filters === undefined) filters = pdf.resolve(streamObj.stream.F);
  let parms = pdf.resolve(streamObj.stream.DecodeParms);
  if (parms === undefined) parms = pdf.resolve(streamObj.stream.DP);

  if (filters === undefined) return data.toString('latin1');
  if (!Array.isArray(filters)) filters = [filters];
  if (parms !== undefined && !Array.isArray(parms)) parms = [parms];

  for (let i = 0; i < filters.length; i++) {
    const f = nameValue(filters[i]);
    const pm = Array.isArray(parms) ? parms[i] : undefined;
    if (f === 'FlateDecode' || f === 'Fl') {
      let out = null;
      try {
        out = zlib.inflateSync(data, { maxOutputLength: INFLATE_CEILING });
      } catch (e1) {
        try {
          out = zlib.inflateRawSync(data, { maxOutputLength: INFLATE_CEILING });
        } catch (e2) {
          return data.toString('latin1');
        }
      }
      const rp = pm ? pdf.resolve(pm) : undefined;
      if (rp && typeof rp === 'object' && !Array.isArray(rp) && rp.Predictor !== undefined) {
        out = applyPredictor(out, rp, pdf);
      }
      data = out;
    } else if (f === 'ASCIIHexDecode' || f === 'AHx') {
      data = Buffer.from(hexToBytes(data.toString('latin1')), 'latin1');
    } else if (f === 'ASCII85Decode' || f === 'A85') {
      const dec = ascii85Decode(data.toString('latin1'));
      if (dec === null) return data.toString('latin1');
      data = dec;
    } else if (f === 'LZWDecode' || f === 'DCTDecode' || f === 'JPXDecode' ||
               f === 'CCITTFaxDecode' || f === 'JBIG2Decode') {
      // Image compression filters: not needed for text, leave the bytes alone.
      return data.toString('latin1');
    } else if (f === 'RunLengthDecode' || f === 'RL') {
      data = runLengthDecode(data);
    } else {
      return data.toString('latin1');
    }
  }
  return data.toString('latin1');
}

function ascii85Decode(str) {
  const out = [];
  let tuple = 0;
  let count = 0;
  for (let i = 0; i < str.length; i++) {
    const c = str[i];
    if (isWS(c)) continue;
    if (c === '~') break;
    if (c === 'z' && count === 0) { out.push(0, 0, 0, 0); continue; }
    const code = str.charCodeAt(i) - 33;
    if (code < 0 || code > 84) return null;
    tuple = tuple * 85 + code;
    count++;
    if (count === 5) {
      out.push((tuple >>> 24) & 0xff, (tuple >>> 16) & 0xff, (tuple >>> 8) & 0xff, tuple & 0xff);
      tuple = 0;
      count = 0;
    }
  }
  if (count > 0) {
    for (let i = count; i < 5; i++) tuple = tuple * 85 + 84;
    const b = [(tuple >>> 24) & 0xff, (tuple >>> 16) & 0xff, (tuple >>> 8) & 0xff, tuple & 0xff];
    for (let i = 0; i < count - 1; i++) out.push(b[i]);
  }
  return Buffer.from(out);
}

function runLengthDecode(buf) {
  const out = [];
  let p = 0;
  while (p < buf.length) {
    const l = buf[p++];
    if (l === 128) break;
    if (l < 128) {
      for (let i = 0; i <= l && p < buf.length; i++) out.push(buf[p++]);
    } else {
      const b = buf[p++];
      for (let i = 0; i < 257 - l; i++) out.push(b);
    }
  }
  return Buffer.from(out);
}

// ---------------------------------------------------------------------------
// PDF document: xref, object map, lazy resolution
// ---------------------------------------------------------------------------

class Pdf {
  constructor(buf) {
    this.raw = buf;
    this.str = buf.toString('latin1');
    this.offsets = new Map(); // objNum -> { type, offset | objstm, index }
    this.cache = new Map();
    this.objStmCache = new Map();
    this.resolving = new Set();
    this.warnings = [];
    this.trailer = null;
    this.xrefOk = false;
    this._parseXref();
    this._scanFallback();
  }

  warn(msg) {
    if (!this.warnings.includes(msg)) this.warnings.push(msg);
  }

  _parseXref() {
    try {
      const m = /startxref\s+(\d+)/g;
      let last = null;
      let hit;
      while ((hit = m.exec(this.str)) !== null) last = hit;
      if (!last) return;
      const start = parseInt(last[1], 10);
      if (!(start >= 0 && start < this.str.length)) return;
      this._readXrefChain(start, new Set());
      this.xrefOk = this.offsets.size > 0;
    } catch (e) {
      this.warn('Cross-reference table could not be read; objects were found by scanning.');
    }
  }

  _readXrefChain(offset, visited) {
    if (visited.has(offset) || visited.size > 64) return;
    visited.add(offset);
    let trailer = null;

    const p = skipWS(this.str, offset);
    if (this.str.startsWith('xref', p)) {
      trailer = this._readClassicXref(p + 4);
    } else {
      const r = parseValue(this.str, p);
      if (r && isStream(r.v)) {
        trailer = this._readXrefStream(r.v);
      }
    }

    if (trailer) {
      if (!this.trailer) this.trailer = {};
      for (const k of Object.keys(trailer)) {
        if (this.trailer[k] === undefined) this.trailer[k] = trailer[k];
      }
      if (typeof trailer.XRefStm === 'number') {
        try {
          const r = parseValue(this.str, trailer.XRefStm);
          if (r && isStream(r.v)) this._readXrefStream(r.v);
        } catch (e) { /* hybrid stream unreadable — the table above is enough */ }
      }
      if (typeof trailer.Prev === 'number') this._readXrefChain(trailer.Prev, visited);
    }
  }

  _readClassicXref(p) {
    const entryRe = /(\d{10})\s+(\d{5})\s+([nf])/y;
    for (;;) {
      p = skipWS(this.str, p);
      if (p >= this.str.length) break;
      if (this.str.startsWith('trailer', p)) {
        const r = parseValue(this.str, p + 7);
        return r && typeof r.v === 'object' ? r.v : null;
      }
      const first = readRegular(this.str, p);
      if (!/^\d+$/.test(first.token)) break;
      p = skipWS(this.str, first.p);
      const countTok = readRegular(this.str, p);
      if (!/^\d+$/.test(countTok.token)) break;
      const start = parseInt(first.token, 10);
      const count = parseInt(countTok.token, 10);
      p = countTok.p;
      for (let i = 0; i < count; i++) {
        p = skipWS(this.str, p);
        entryRe.lastIndex = p;
        const mm = entryRe.exec(this.str);
        if (!mm) { i = count; break; }
        p = entryRe.lastIndex;
        const num = start + i;
        const off = parseInt(mm[1], 10);
        if (mm[3] === 'n' && !this.offsets.has(num)) {
          this.offsets.set(num, { type: 1, offset: off });
        }
      }
      // Note: a subsection with count 0 is legal (an empty incremental-update
      // xref). Do NOT break here — the following "trailer" still has to be read
      // so that /Prev can be followed to the real table.
    }
    return null;
  }

  _readXrefStream(streamObj) {
    const dict = streamObj.stream;
    const data = Buffer.from(decodeStream(this, streamObj), 'latin1');
    const w = (dict.W || []).map((x) => numberOr(this.resolve(x), 0));
    if (w.length < 3) return dict;
    const size = numberOr(this.resolve(dict.Size), 0);
    let index = this.resolve(dict.Index);
    if (!Array.isArray(index) || index.length < 2) index = [0, size];
    let p = 0;
    const readField = (n) => {
      let v = 0;
      for (let i = 0; i < n; i++) v = v * 256 + (p < data.length ? data[p++] : 0);
      return v;
    };
    for (let s = 0; s + 1 < index.length; s += 2) {
      let start = numberOr(this.resolve(index[s]), 0);
      const count = numberOr(this.resolve(index[s + 1]), 0);
      for (let i = 0; i < count; i++, start++) {
        const t = w[0] === 0 ? 1 : readField(w[0]);
        const f2 = readField(w[1]);
        const f3 = readField(w[2]);
        if (t === 1 && !this.offsets.has(start)) {
          this.offsets.set(start, { type: 1, offset: f2 });
        } else if (t === 2 && !this.offsets.has(start)) {
          this.offsets.set(start, { type: 2, objstm: f2, index: f3 });
        }
      }
    }
    return dict;
  }

  /** Fill gaps with a crude raw scan; xref entries always win. */
  _scanFallback() {
    const re = /(\d{1,10})\s+(\d{1,5})\s+obj\b/g;
    let m;
    let found = 0;
    while ((m = re.exec(this.str)) !== null) {
      const num = parseInt(m[1], 10);
      if (this.offsets.has(num)) continue;
      // Guard against "12 0 obj" text embedded inside a stream: require the
      // match to be followed by something object-ish fairly soon.
      const after = skipWS(this.str, re.lastIndex);
      const c = this.str[after];
      if (c === '<' || c === '[' || c === '/' || c === '(' || c === '-' ||
          c === '+' || c === '.' || (c >= '0' && c <= '9') ||
          this.str.startsWith('true', after) || this.str.startsWith('false', after) ||
          this.str.startsWith('null', after)) {
        this.offsets.set(num, { type: 1, offset: m.index });
        found++;
        if (found > 200000) break;
      }
    }
  }

  /** Dereference a value; refs are followed lazily. */
  resolve(v) {
    let cur = v;
    let guard = 0;
    while (isRef(cur) && guard++ < 64) {
      cur = this.getObject(cur.ref);
    }
    return cur;
  }

  getObject(num) {
    if (num === undefined || num === null) return undefined;
    if (this.cache.has(num)) return this.cache.get(num);
    const entry = this.offsets.get(num);
    if (!entry) return undefined;
    if (this.resolving.has(num)) return undefined;
    this.resolving.add(num);
    let val;
    try {
      if (entry.type === 2) {
        val = this._getFromObjStm(entry.objstm, entry.index, num);
      } else if (entry.type === 1) {
        val = this._parseIndirectAt(entry.offset, num);
      }
    } catch (e) {
      val = undefined;
    }
    this.resolving.delete(num);
    this.cache.set(num, val);
    return val;
  }

  _parseIndirectAt(offset, expectNum) {
    let p = skipWS(this.str, offset);
    const n1 = readRegular(this.str, p);
    if (!/^\d+$/.test(n1.token)) return undefined;
    p = skipWS(this.str, n1.p);
    const n2 = readRegular(this.str, p);
    if (!/^\d+$/.test(n2.token)) return undefined;
    p = skipWS(this.str, n2.p);
    const kw = readRegular(this.str, p);
    if (kw.token !== 'obj') return undefined;
    const r = parseValue(this.str, kw.p);
    if (!r) return undefined;
    let v = r.v;
    const after = skipWS(this.str, r.p);
    if (this.str.startsWith('stream', after)) {
      const data = this._readStreamData(after + 6, v);
      if (data !== null) v = { stream: v, data };
    }
    if (v && typeof v === 'object' && !isName(v) && !isRef(v) && !isPdfString(v) && !Array.isArray(v) && !isStream(v)) {
      v.__num = parseInt(n1.token, 10);
    }
    return v;
  }

  _readStreamData(p, dict) {
    if (this.str[p] === '\r') p++;
    if (this.str[p] === '\n') p++;
    let len = this.resolve(dict && dict.Length);
    if (typeof len === 'number' && len >= 0 && p + len <= this.str.length) {
      const end = this.str.indexOf('endstream', p + len);
      if (end >= 0 && end - (p + len) < 8) return this.str.slice(p, p + len);
    }
    const end = this.str.indexOf('endstream', p);
    if (end < 0) return null;
    let e = end;
    if (e > p && this.str[e - 1] === '\n') e--;
    if (e > p && this.str[e - 1] === '\r') e--;
    return this.str.slice(p, e);
  }

  _getFromObjStm(objstmNum, index, wantNum) {
    let table = this.objStmCache.get(objstmNum);
    if (!table) {
      const stmObj = this.getObject(objstmNum);
      if (!isStream(stmObj)) return undefined;
      const decoded = decodeStream(this, stmObj);
      const n = numberOr(this.resolve(stmObj.stream.N), 0);
      const first = numberOr(this.resolve(stmObj.stream.First), 0);
      const pairs = [];
      let p = 0;
      for (let i = 0; i < n; i++) {
        p = skipWS(decoded, p);
        const a = readRegular(decoded, p);
        if (!/^\d+$/.test(a.token)) break;
        p = skipWS(decoded, a.p);
        const b = readRegular(decoded, p);
        if (!/^\d+$/.test(b.token)) break;
        p = b.p;
        pairs.push({ num: parseInt(a.token, 10), off: parseInt(b.token, 10) });
      }
      table = { decoded, first, pairs };
      this.objStmCache.set(objstmNum, table);
    }
    let pair = table.pairs[index];
    if (!pair || pair.num !== wantNum) pair = table.pairs.find((x) => x.num === wantNum);
    if (!pair) return undefined;
    const r = parseValue(table.decoded, table.first + pair.off);
    return r ? r.v : undefined;
  }

  /** The document catalog. */
  catalog() {
    if (this.trailer) {
      const root = this.resolve(this.trailer.Root);
      if (root && typeof root === 'object' && !Array.isArray(root)) return root;
    }
    for (const num of Array.from(this.offsets.keys()).sort((a, b) => a - b)) {
      const o = this.getObject(num);
      if (o && typeof o === 'object' && nameValue(this.resolve(o.Type)) === 'Catalog') return o;
    }
    return null;
  }
}

// ---------------------------------------------------------------------------
// CMap / ToUnicode parsing
// ---------------------------------------------------------------------------

function tokenizeCMapBlock(block) {
  const items = [];
  const re = /\/[^\s<>\[\]()]*|<[0-9A-Fa-f\s]*>|\[|\]/g;
  let m;
  let array = null;
  while ((m = re.exec(block)) !== null) {
    const t = m[0];
    if (t === '[') { array = []; items.push({ t: 'array', items: array }); continue; }
    if (t === ']') { array = null; continue; }
    if (t[0] === '/') {
      const name = t.slice(1);
      if (array) array.push({ t: 'name', v: name });
      else items.push({ t: 'name', v: name });
      continue;
    }
    const hex = t.slice(1, -1);
    const item = { t: 'hex', v: hex };
    if (array) array.push(item);
    else items.push(item);
  }
  return items;
}

/** Parse a ToUnicode CMap into a code -> unicode string map plus code lengths. */
function parseCMap(text) {
  const map = new Map();
  const rangeLens = new Set();

  const section = (name) => {
    const re = new RegExp('begin' + name + '([\\s\\S]*?)end' + name, 'g');
    const out = [];
    let m;
    while ((m = re.exec(text)) !== null) out.push(m[1]);
    return out;
  };

  for (const block of section('codespacerange')) {
    const items = tokenizeCMapBlock(block).filter((x) => x.t === 'hex');
    for (let i = 0; i + 1 < items.length; i += 2) {
      const len = Math.max(1, Math.ceil(items[i].v.replace(/[^0-9A-Fa-f]/g, '').length / 2));
      rangeLens.add(len);
    }
  }

  for (const block of section('bfchar')) {
    const items = tokenizeCMapBlock(block);
    for (let i = 0; i + 1 < items.length; i += 2) {
      const src = items[i];
      const dst = items[i + 1];
      if (src.t !== 'hex') continue;
      const srcHex = src.v.replace(/[^0-9A-Fa-f]/g, '');
      const code = parseInt(srcHex || '0', 16);
      rangeLens.add(Math.max(1, Math.ceil(srcHex.length / 2)));
      let u = '';
      if (dst.t === 'hex') u = hexToUtf16(dst.v);
      else if (dst.t === 'name') u = dst.v === 'space' ? ' ' : '';
      if (u) map.set(code, u);
    }
  }

  for (const block of section('bfrange')) {
    const items = tokenizeCMapBlock(block);
    for (let i = 0; i + 2 < items.length; i += 3) {
      const lo = items[i];
      const hi = items[i + 1];
      const dst = items[i + 2];
      if (lo.t !== 'hex' || hi.t !== 'hex') continue;
      const loHex = lo.v.replace(/[^0-9A-Fa-f]/g, '');
      const hiHex = hi.v.replace(/[^0-9A-Fa-f]/g, '');
      const loCode = parseInt(loHex || '0', 16);
      const hiCode = parseInt(hiHex || '0', 16);
      rangeLens.add(Math.max(1, Math.ceil(loHex.length / 2)));
      if (dst.t === 'array') {
        const arr = dst.items || [];
        for (let k = 0; loCode + k <= hiCode && k < arr.length; k++) {
          if (arr[k].t !== 'hex') continue;
          const u = hexToUtf16(arr[k].v);
          if (u) map.set(loCode + k, u);
        }
      } else if (dst.t === 'hex') {
        const units = [];
        const clean = dst.v.replace(/[^0-9A-Fa-f]/g, '');
        for (let k = 0; k + 4 <= clean.length; k += 4) {
          units.push(parseInt(clean.substr(k, 4), 16));
        }
        if (!units.length && clean.length) units.push(parseInt(clean, 16));
        if (!units.length) continue;
        for (let code = loCode; code <= hiCode; code++) {
          const step = code - loCode;
          if (units.length === 1) {
            map.set(code, String.fromCharCode((units[0] + step) & 0xffff));
          } else {
            const copy = units.slice();
            copy[copy.length - 1] = (copy[copy.length - 1] + step) & 0xffff;
            map.set(code, String.fromCharCode.apply(null, copy));
          }
        }
      }
    }
  }

  // Fall back to Inferring code length from the mapped values.
  if (!rangeLens.size && map.size) {
    let max = 0;
    for (const k of map.keys()) if (k > max) max = k;
    rangeLens.add(max > 0xff ? 2 : 1);
  }

  return { map, lens: Array.from(rangeLens).sort((a, b) => b - a) };
}

// WinAnsi (cp1252) high range, for fonts with no ToUnicode map.
const CP1252_HIGH = {
  0x80: '\u20AC', 0x82: '\u201A', 0x83: '\u0192', 0x84: '\u201E', 0x85: '\u2026',
  0x86: '\u2020', 0x87: '\u2021', 0x88: '\u02C6', 0x89: '\u2030', 0x8a: '\u0160',
  0x8b: '\u2039', 0x8c: '\u0152', 0x8e: '\u017D', 0x91: '\u2018', 0x92: '\u2019',
  0x93: '\u201C', 0x94: '\u201D', 0x95: '\u2022', 0x96: '\u2013', 0x97: '\u2014',
  0x98: '\u02DC', 0x99: '\u2122', 0x9a: '\u0161', 0x9b: '\u203A', 0x9c: '\u0153',
  0x9e: '\u017E', 0x9f: '\u0178',
};

function fallbackByte(b) {
  if (b === 0x20) return ' ';
  if (b >= 0x20 && b < 0x7f) return String.fromCharCode(b);
  if (CP1252_HIGH[b]) return CP1252_HIGH[b];
  if (b >= 0xa0) return String.fromCharCode(b);
  return '';
}

/** Build a per-font byte decoder from the font dictionary. */
function buildFontDecoder(pdf, fontObj) {
  let cmap = null;
  if (fontObj && typeof fontObj === 'object') {
    const tu = pdf.resolve(fontObj.ToUnicode);
    if (isStream(tu)) {
      try {
        const parsed = parseCMap(decodeStream(pdf, tu));
        if (parsed.map.size) cmap = parsed;
      } catch (e) { /* fall through to the encoding-based path */ }
    }
  }
  if (cmap) {
    const map = cmap.map;
    const lens = cmap.lens.length ? cmap.lens : [2];
    return function decode(str) {
      const bytes = [];
      for (let i = 0; i < str.length; i++) bytes.push(str.charCodeAt(i) & 0xff);
      let out = '';
      let i = 0;
      while (i < bytes.length) {
        let done = false;
        for (let li = 0; li < lens.length; li++) {
          const L = lens[li];
          if (i + L > bytes.length) continue;
          let code = 0;
          for (let k = 0; k < L; k++) code = code * 256 + bytes[i + k];
          const u = map.get(code);
          if (u !== undefined) { out += u; i += L; done = true; break; }
        }
        if (!done) { out += fallbackByte(bytes[i]); i++; }
      }
      return out;
    };
  }
  // No usable ToUnicode: treat as a simple single-byte font.
  return function decode(str) {
    let out = '';
    for (let i = 0; i < str.length; i++) out += fallbackByte(str.charCodeAt(i) & 0xff);
    return out;
  };
}

// ---------------------------------------------------------------------------
// Content stream text extraction
// ---------------------------------------------------------------------------

/** Read the next token from a content stream. */
function nextContentToken(s, p) {
  p = skipWS(s, p);
  if (p >= s.length) return { type: 'eof', p };
  const c = s[p];

  if (c === '<') {
    if (s[p + 1] === '<') {
      const r = parseDict(s, p);
      return { type: 'dict', value: r.v, p: r.p };
    }
    const end = s.indexOf('>', p + 1);
    if (end < 0) return { type: 'str', value: '', p: s.length };
    return { type: 'str', value: hexToBytes(s.slice(p + 1, end)), p: end + 1 };
  }
  if (c === '(') {
    const r = parseLiteralString(s, p);
    return { type: 'str', value: r.v.str, p: r.p };
  }
  if (c === '[') {
    p++;
    const arr = [];
    for (;;) {
      p = skipWS(s, p);
      if (p >= s.length) break;
      if (s[p] === ']') { p++; break; }
      const t = nextContentToken(s, p);
      if (t.type === 'eof') break;
      if (t.p === p) { p++; continue; }
      arr.push(t.type === 'op' ? { name: t.value } : t.value);
      p = t.p;
    }
    return { type: 'array', value: arr, p };
  }
  if (c === '/') {
    const r = readRegular(s, p + 1);
    return { type: 'name', value: r.token, p: r.p };
  }
  if (c === '+' || c === '-' || c === '.' || (c >= '0' && c <= '9')) {
    const r = readRegular(s, p);
    return { type: 'num', value: Number(r.token), p: r.p };
  }
  const r = readRegular(s, p);
  if (!r.token) return { type: 'eof', p: p + 1 };
  return { type: 'op', value: r.token, p: r.p };
}

/**
 * Walk one content stream applying a small text state machine. Text pieces and
 * explicit separators are pushed into `out`; cleanup happens once at the end.
 */
function runContentStream(pdf, content, pageCtx, out, depth) {
  if (depth > 8) return;
  const operands = [];
  const decoders = pageCtx.fontDecoders;
  let decoder = null;
  let p = 0;

  const pushText = (bytes) => {
    const text = decoder ? decoder(bytes) : defaultDecode(bytes);
    if (text) out.push(text);
  };
  const newline = () => {
    const last = out.length ? out[out.length - 1] : null;
    if (last !== '\n' && last !== null) out.push('\n');
    else if (last === null) out.push('\n');
  };
  const space = () => {
    const last = out.length ? out[out.length - 1] : null;
    if (last !== null && last !== '\n' && last !== ' ') out.push(' ');
  };

  let lastTmX = null;
  let lastTmY = null;

  while (p < content.length) {
    const t = nextContentToken(content, p);
    if (t.type === 'eof') break;
    p = t.p;

    if (t.type !== 'op') { operands.push(t); continue; }

    const op = t.value;
    operands.push({ type: 'op', value: op });

    switch (op) {
      case 'BT':
        operands.length = 0;
        lastTmX = null;
        lastTmY = null;
        break;
      case 'ET':
        newline();
        operands.length = 0;
        break;
      case 'Tf': {
        const nm = findName(operands);
        if (nm) decoder = decoders.get(nm) || null;
        operands.length = 0;
        break;
      }
      case 'Td': case 'TD': {
        newline();
        operands.length = 0;
        break;
      }
      case 'T*':
        newline();
        operands.length = 0;
        break;
      case 'TL': case 'Tc': case 'Tw': case 'Tz': case 'Ts': case 'Tr':
        operands.length = 0;
        break;
      case 'Tm': {
        const nums = operands.filter((o) => o && o.type === 'num').map((o) => o.value);
        if (nums.length >= 6) {
          const x = nums[4];
          const y = nums[5];
          if (lastTmY === null || Math.abs(y - lastTmY) > 0.5) newline();
          else if (lastTmX !== null && Math.abs(x - lastTmX) > 6) space();
          lastTmX = x;
          lastTmY = y;
        }
        operands.length = 0;
        break;
      }
      case 'cm':
        operands.length = 0;
        break;
      case 'Tj': {
        const s = findString(operands);
        if (s !== null) pushText(s);
        operands.length = 0;
        break;
      }
      case "'": {
        newline();
        const s = findString(operands);
        if (s !== null) pushText(s);
        operands.length = 0;
        break;
      }
      case '"': {
        newline();
        const s = findString(operands);
        if (s !== null) pushText(s);
        operands.length = 0;
        break;
      }
      case 'TJ': {
        const arr = findArray(operands);
        if (arr) {
          for (const item of arr) {
            if (typeof item === 'string') pushText(item);
            else if (isPdfString(item)) pushText(item.str);
            else if (typeof item === 'number' && item < -100) space();
          }
        }
        operands.length = 0;
        break;
      }
      case 'Do': {
        const nm = findName(operands);
        if (nm && pageCtx.xobjects) {
          const xo = pdf.resolve(pageCtx.xobjects[nm]);
          if (isStream(xo) && nameValue(pdf.resolve(xo.stream.Subtype)) === 'Form') {
            const sub = {
              fontDecoders: buildFontDecoders(pdf, pdf.resolve(xo.stream.Resources)) || pageCtx.fontDecoders,
              xobjects: xobjectMap(pdf, pdf.resolve(xo.stream.Resources)) || pageCtx.xobjects,
            };
            runContentStream(pdf, decodeStream(pdf, xo), sub, out, depth + 1);
          }
        }
        operands.length = 0;
        break;
      }
      case 'BI': {
        // Inline image: skip everything up to the matching EI.
        const idIdx = content.indexOf('ID', p);
        if (idIdx < 0) { p = content.length; break; }
        let q = idIdx + 2;
        if (isWS(content[q])) q++;
        const ei = findInlineImageEnd(content, q);
        p = ei < 0 ? content.length : ei + 2;
        operands.length = 0;
        break;
      }
      default:
        operands.length = 0;
    }
  }
}

function findInlineImageEnd(s, from) {
  for (let i = from; i < s.length - 1; i++) {
    if (s[i] === 'E' && s[i + 1] === 'I') {
      const before = s[i - 1];
      const after = s[i + 2];
      if ((before === undefined || isWS(before)) && (after === undefined || isWS(after))) return i;
    }
  }
  return -1;
}

function defaultDecode(str) {
  let out = '';
  for (let i = 0; i < str.length; i++) out += fallbackByte(str.charCodeAt(i) & 0xff);
  return out;
}

function findName(operands) {
  for (let i = operands.length - 1; i >= 0; i--) if (operands[i] && operands[i].type === 'name') return operands[i].value;
  return null;
}
function findString(operands) {
  for (let i = operands.length - 1; i >= 0; i--) if (operands[i] && operands[i].type === 'str') return operands[i].value;
  return null;
}
function findArray(operands) {
  for (let i = operands.length - 1; i >= 0; i--) if (operands[i] && operands[i].type === 'array') return operands[i].value;
  return null;
}

/** Map font resource names to decoders for one page. */
function buildFontDecoders(pdf, resources) {
  const map = new Map();
  const res = pdf.resolve(resources);
  if (!res || typeof res !== 'object') return map;
  const fonts = pdf.resolve(res.Font);
  if (!fonts || typeof fonts !== 'object') return map;
  for (const key of Object.keys(fonts)) {
    const f = pdf.resolve(fonts[key]);
    if (f && typeof f === 'object') {
      try { map.set(key, buildFontDecoder(pdf, f)); } catch (e) { /* fall back to default */ }
    }
  }
  return map;
}

function xobjectMap(pdf, resources) {
  const res = pdf.resolve(resources);
  if (!res || typeof res !== 'object') return null;
  const xo = pdf.resolve(res.XObject);
  if (!xo || typeof xo !== 'object') return null;
  return xo;
}

// ---------------------------------------------------------------------------
// Page tree
// ---------------------------------------------------------------------------

function collectPages(pdf) {
  const pages = [];
  const cat = pdf.catalog();
  const seen = new Set();
  let ok = false;
  if (cat) {
    const root = pdf.resolve(cat.Pages);
    if (root && typeof root === 'object') {
      ok = walk(root, undefined, 0);
    }
  }
  function walk(node, inherited, depth) {
    if (depth > 64 || seen.has(node)) return false;
    seen.add(node);
    const type = nameValue(pdf.resolve(node.Type));
    const resources = node.Resources !== undefined ? node.Resources : inherited;
    if (type === 'Pages' || (node.Kids !== undefined && type !== 'Page')) {
      const kids = pdf.resolve(node.Kids);
      if (!Array.isArray(kids)) return false;
      for (const k of kids) {
        const kn = pdf.resolve(k);
        if (kn && typeof kn === 'object') {
          if (!walk(kn, resources, depth + 1) && nameValue(pdf.resolve(kn.Type)) === 'Page') {
            pages.push({ dict: kn, resources: kn.Resources !== undefined ? kn.Resources : resources });
          }
        }
      }
      return true;
    }
    pages.push({ dict: node, resources });
    return true;
  }

  if (pages.length) return { pages, fromTree: true };

  // Fallback: object-number order.
  const nums = Array.from(pdf.offsets.keys()).sort((a, b) => a - b);
  for (const num of nums) {
    const o = pdf.getObject(num);
    if (o && typeof o === 'object' && !Array.isArray(o) && nameValue(pdf.resolve(o.Type)) === 'Page') {
      pages.push({ dict: o, resources: o.Resources });
    }
  }
  return { pages, fromTree: false };
}

// ---------------------------------------------------------------------------
// Per-page text
// ---------------------------------------------------------------------------

function pageContentString(pdf, page) {
  const contents = pdf.resolve(page.Contents);
  if (contents === undefined || contents === null) return '';
  const parts = Array.isArray(contents) ? contents : [contents];
  let out = '';
  for (const c of parts) {
    const s = pdf.resolve(c);
    if (isStream(s)) {
      try { out += decodeStream(pdf, s) + '\n'; } catch (e) { /* skip unreadable part */ }
    }
  }
  return out;
}

function pageHasImages(pdf, resources) {
  const xo = xobjectMap(pdf, resources);
  if (!xo) return false;
  for (const key of Object.keys(xo)) {
    const x = pdf.resolve(xo[key]);
    if (isStream(x) && nameValue(pdf.resolve(x.stream.Subtype)) === 'Image') return true;
  }
  return false;
}

function cleanText(raw) {
  let t = raw.replace(/\r\n?/g, '\n');
  t = t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');
  t = t.replace(/\u00A0/g, ' ');
  const lines = t.split('\n').map((l) => l.replace(/[ \t]+/g, ' ').trim());
  const kept = lines.filter((l) => l.length > 0);
  return kept.join('\n');
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Extract the text of every page from a PDF buffer.
 *
 * @param {Buffer} buffer
 * @returns {{pageCount:number, pages:Array<{index:number,text:string,hasImages:boolean}>, warnings:string[]}}
 * Never throws: unreadable input yields an empty result plus a warning.
 */
function extractPdfText(buffer) {
  const warnings = [];
  const empty = { pageCount: 0, pages: [], warnings };
  try {
    if (!buffer || !buffer.length) {
      warnings.push('The PDF was empty.');
      return empty;
    }
    const head = buffer.toString('latin1', 0, Math.min(1024, buffer.length));
    if (!head.includes('%PDF-')) {
      warnings.push('The file does not look like a PDF (no %PDF- header).');
    }
    const pdf = new Pdf(buffer);
    if (!pdf.xrefOk) {
      pdf.warn('The cross-reference table was missing or damaged; objects were located by scanning the file.');
    }
    const { pages, fromTree } = collectPages(pdf);
    if (!fromTree && pages.length) {
      pdf.warn('The page tree could not be resolved; pages are listed in object order.');
    }
    if (!pages.length) {
      pdf.warn('No pages could be found in this PDF.');
    }
    const outPages = [];
    for (let i = 0; i < pages.length; i++) {
      const pg = pages[i];
      let text = '';
      let hasImages = false;
      try {
        hasImages = pageHasImages(pdf, pg.resources);
      } catch (e) { hasImages = false; }
      try {
        const decoders = buildFontDecoders(pdf, pg.resources);
        const content = pageContentString(pdf, pg.dict);
        const out = [];
        runContentStream(pdf, content, {
          fontDecoders: decoders,
          xobjects: xobjectMap(pdf, pg.resources),
        }, out, 0);
        text = cleanText(out.join(''));
      } catch (e) {
        pdf.warn('One page could not be read (' + (e && e.message ? e.message : 'unknown error') + ').');
      }
      outPages.push({ index: i + 1, text, hasImages });
    }
    if (outPages.length && outPages.every((p) => !p.text.length)) {
      pdf.warn('No text could be extracted; the PDF may contain scanned images rather than text.');
    }
    for (const w of pdf.warnings) warnings.push(w);
    return { pageCount: outPages.length, pages: outPages, warnings };
  } catch (e) {
    warnings.push('The PDF could not be read (' + (e && e.message ? e.message : 'unknown error') + ').');
    return empty;
  }
}

/** Convenience wrapper: read a file from disk and extract its text. */
function extractPdfTextFile(path) {
  try {
    const buf = fs.readFileSync(path);
    return extractPdfText(buf);
  } catch (e) {
    return {
      pageCount: 0,
      pages: [],
      warnings: ['The file could not be opened (' + (e && e.message ? e.message : 'unknown error') + ').'],
    };
  }
}

module.exports = { extractPdfText, extractPdfTextFile };
