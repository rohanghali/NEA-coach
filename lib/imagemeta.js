'use strict';
/*
 * imagemeta.js — read pixel dimensions out of image files without decoding them.
 *
 * We only need width/height, so we read the file headers. That is instant and
 * needs no image library. Supports PNG, JPEG, GIF and BMP, which covers
 * essentially everything PowerPoint embeds.
 */

/** PNG: 8-byte signature, then an IHDR chunk holding width and height. */
function png(buf) {
  if (buf.length < 24) return null;
  if (buf.readUInt32BE(0) !== 0x89504e47) return null;
  return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20), type: 'png' };
}

/** GIF: 'GIF87a'/'GIF89a', then logical screen width/height as little-endian u16. */
function gif(buf) {
  if (buf.length < 10) return null;
  const sig = buf.toString('ascii', 0, 3);
  if (sig !== 'GIF') return null;
  return { width: buf.readUInt16LE(6), height: buf.readUInt16LE(8), type: 'gif' };
}

/** BMP: 'BM', then a DIB header with width/height as signed little-endian i32. */
function bmp(buf) {
  if (buf.length < 26) return null;
  if (buf.toString('ascii', 0, 2) !== 'BM') return null;
  return {
    width: Math.abs(buf.readInt32LE(18)),
    height: Math.abs(buf.readInt32LE(22)),
    type: 'bmp',
  };
}

/**
 * JPEG: walk the marker segments until we hit a Start-Of-Frame, which carries
 * the dimensions. We skip entropy-coded data carefully rather than scanning
 * blindly, so we do not mistake image bytes for markers.
 */
function jpeg(buf) {
  if (buf.length < 4) return null;
  if (buf[0] !== 0xff || buf[1] !== 0xd8) return null; // SOI

  let p = 2;
  while (p + 3 < buf.length) {
    if (buf[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = buf[p + 1];

    // Padding and standalone markers carry no length field.
    if (marker === 0xff || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
      p += 2;
      continue;
    }

    const len = buf.readUInt16BE(p + 2);
    if (len < 2) return null;

    const isSOF =
      (marker >= 0xc0 && marker <= 0xc3) ||
      (marker >= 0xc5 && marker <= 0xc7) ||
      (marker >= 0xc9 && marker <= 0xcb) ||
      (marker >= 0xcd && marker <= 0xcf);

    if (isSOF && p + 9 < buf.length) {
      return {
        height: buf.readUInt16BE(p + 5),
        width: buf.readUInt16BE(p + 7),
        type: 'jpeg',
      };
    }

    p += 2 + len;
  }
  return null;
}

/** Detect the format from the header bytes and return {width, height, type}. */
function getDimensions(buf) {
  if (!buf || buf.length < 8) return null;
  return png(buf) || jpeg(buf) || gif(buf) || bmp(buf) || null;
}

/**
 * Find the pixel dimensions of an EMF/WMF vector image from its header.
 * Vector images scale cleanly, so we mark them as such rather than guessing
 * a resolution and falsely flagging them as "low quality".
 */
function getVectorInfo(buf) {
  if (!buf || buf.length < 44) return null;

  // EMF: record type 1 starts with 0x00000001 and has bounds in the header.
  if (buf.readUInt32LE(0) === 0x00000001) {
    const type = buf.readUInt32LE(0);
    const sig = buf.readUInt32LE(40);
    if (type === 1 && sig === 0x464d4520) {
      // ' EMF'
      const left = buf.readInt32LE(8);
      const top = buf.readInt32LE(12);
      const right = buf.readInt32LE(16);
      const bottom = buf.readInt32LE(20);
      return { type: 'emf', width: Math.abs(right - left), height: Math.abs(bottom - top) };
    }
  }

  // WMF: placeable header magic 0x9AC6CDD7
  if (buf.readUInt32LE(0) === 0x9ac6cdd7) {
    return { type: 'wmf', width: null, height: null };
  }

  return null;
}

module.exports = { getDimensions, getVectorInfo };
