/*
 * NEA Coach — the app icon, drawn from one description and written out in every
 * format that needs it.
 *
 *   node tools/make-icons.js
 *
 * WHY THIS EXISTS
 *
 * There was no icon anywhere: not in the browser tab, not on a phone's home screen
 * (which is what the student reported — "i have it on my home screen on mobile but
 * theres no icon"), and not on the desktop. The app is distributed as a folder, so
 * there was nowhere for one to come from.
 *
 * WHY IT IS WRITTEN BY HAND
 *
 * The project ships with zero third-party dependencies and no build step, and an icon
 * does not justify breaking that. A PNG is a header, a zlib stream and two checksums,
 * all of which Node has built in, and ICO and ICNS are containers of PNGs. So the whole
 * thing is a rasteriser, a PNG encoder and two thin wrappers.
 *
 * ONE DRAWING, MANY SIZES
 *
 * Everything below paints a single function of (x, y) in 0..1.
 *
 * The mark is three flowing bands — the same drapes that move behind the app — so the icon and
 * the app are recognisably one thing. The favicon, the phone
 * icon, the Windows icon and the Mac icon are all THAT drawing sampled at different
 * sizes and, in two cases, with different framing — so they cannot drift apart into
 * four subtly different logos.
 *
 * The two framings are not fussiness:
 *
 *   - iOS applies its own rounded mask to a home-screen icon, so it wants a FULL-BLEED
 *     square. Insetting it as well produces a rounded square inside a rounded mask,
 *     which is visibly wrong at the corners.
 *   - macOS does NOT mask an .icns, so a full-bleed square looks like a sticker. It
 *     wants the artwork inset with a shadow, the way every other Mac icon is drawn.
 *
 * Outputs are committed, so a fresh clone or a packaged build needs nothing run. This
 * only matters when the icon itself changes.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const APP_DIR = path.join(__dirname, '..');
const PUBLIC_DIR = path.join(APP_DIR, 'public');
const ASSET_DIR = path.join(APP_DIR, 'assets');

/* ------------------------------------------------------------------ *
 * The drawing
 * ------------------------------------------------------------------ */

const ACCENT = [0x25, 0x63, 0xeb];      // the app's own accent
const BADGE_TOP = [0x5b, 0x8c, 0xf8];
const BADGE_BOTTOM = [0x1a, 0x3f, 0xc4];

/** Rounded-rectangle signed distance. Negative inside. */
function roundedRectDist(x, y, x0, y0, x1, y1, r) {
  const cx = (x0 + x1) / 2;
  const cy = (y0 + y1) / 2;
  const hw = Math.max(0, (x1 - x0) / 2 - r);
  const hh = Math.max(0, (y1 - y0) / 2 - r);
  const dx = Math.abs(x - cx) - hw;
  const dy = Math.abs(y - cy) - hh;
  const ax = Math.max(dx, 0);
  const ay = Math.max(dy, 0);
  return Math.sqrt(ax * ax + ay * ay) + Math.min(Math.max(dx, dy), 0) - r;
}

/** Distance from a point to a line segment. Used to stroke the tick. */
function segmentDist(px, py, ax, ay, bx, by) {
  const vx = bx - ax;
  const vy = by - ay;
  const wx = px - ax;
  const wy = py - ay;
  const len2 = vx * vx + vy * vy;
  const t = len2 === 0 ? 0 : Math.max(0, Math.min(1, (wx * vx + wy * vy) / len2));
  const dx = wx - vx * t;
  const dy = wy - vy * t;
  return Math.sqrt(dx * dx + dy * dy);
}

const clamp01 = (n) => (n < 0 ? 0 : n > 1 ? 1 : n);
/** 0 at `edge`, 1 at `edge - width`, smoothly. For soft shadows. */
const falloff = (d, width) => clamp01(1 - d / width);

function mix(a, b, t) {
  return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t];
}

/** src over dst, both [r, g, b, a] with a in 0..1. */
function over(dst, src) {
  const a = src[3] + dst[3] * (1 - src[3]);
  if (a <= 0) { return [0, 0, 0, 0]; }
  const out = [];
  for (let i = 0; i < 3; i++) {
    out[i] = (src[i] * src[3] + dst[i] * dst[3] * (1 - src[3])) / a;
  }
  out[3] = a;
  return out;
}

/*
 * The whole icon, as a function of a point.
 *
 * `inset` is the margin as a fraction of the canvas, and `shadow` adds the drop shadow
 * macOS expects. Everything inside is expressed as fractions of the BADGE, not of the
 * canvas, so insetting the artwork does not silently rescale the design.
 */
function sample(u, v, inset, shadow, maskable, size, square) {
  const span = 1 - inset * 2;
  const bx = (u - inset) / span;      // 0..1 across the badge
  const by = (v - inset) / span;
  /*
   * 22% corners: round enough to read as an app icon, square enough not to look like a blob
   * at 16px. A MASKABLE icon gets square corners on purpose — Android crops it to its own
   * shape, and a rounded source leaves transparent slivers at the extreme corners that the
   * crop does not always cover.
   */
  const radius = (maskable || square) ? 0 : 0.215;

  let out = [0, 0, 0, 0];

  if (shadow) {
    // A soft shadow under the badge, offset down. Two passes, tight and broad, which is
    // what makes it read as a lifted surface rather than as a smear.
    const tight = roundedRectDist(u, v, inset - 0.004, inset + 0.016, 1 - inset + 0.004, 1 - inset + 0.020, radius * span);
    const broad = roundedRectDist(u, v, inset - 0.010, inset + 0.026, 1 - inset + 0.010, 1 - inset + 0.038, radius * span);
    const a = Math.max(0.20 * falloff(tight, 0.020), 0.16 * falloff(broad, 0.055));
    out = over(out, [0x0b, 0x14, 0x2e, a]);
  }

  // The badge.
  const badge = roundedRectDist(bx, by, 0, 0, 1, 1, radius);
  if (badge < 0) {
    let colour = mix(BADGE_TOP, BADGE_BOTTOM, clamp01((bx + by) / 2));
    // A soft light from the top-left, so the gradient does not read as a flat ramp.
    const hl = clamp01(1 - Math.sqrt((bx - 0.22) ** 2 + (by - 0.10) ** 2) / 1.05);
    colour = mix(colour, [255, 255, 255], 0.13 * hl * hl);
    /*
     * A lit rim just inside the upper-left edge — the same "glass edge" the app's own cards
     * carry. Cheap, and it is what separates a drawn icon from a flat rounded rectangle.
     */
    const rim = clamp01(1 + badge / 0.022) * clamp01(1 - Math.sqrt(bx * bx + by * by) / 1.25);
    colour = mix(colour, [255, 255, 255], 0.32 * rim * rim);
    out = over(out, [colour[0], colour[1], colour[2], 1]);
  }
  if (out[3] === 0) { return out; }

  /*
   * The mark: flowing bands, the app's own visual signature.
   *
   * What was here was a white document with a blue tick on it — the generic "checked file"
   * glyph, and the reason the icon read as basic. Anyone can draw a page with a tick; it says
   * nothing about what this is.
   *
   * These are the same layered drapes that move behind the page. The icon and the app now look
   * like the same thing, which is the whole job of a mark at this size.
   *
   * Each band is a signed distance from its own centre line — a sine of x — so the edge can be
   * softened the same way the badge is. A hard threshold would look fine at 512px and ragged at
   * 16, where the whole band is under two pixels tall.
   */
  /*
   * TWO bands when the icon is small, three when it is large.
   *
   * At 16px the badge is sixteen pixels across, so three bands of a pixel and a half — with soft
   * edges on top of that — merge into a blue smudge with a pale streak through it. Rendered at 32
   * and looked at, which is how this was caught: it read as a wave at 512 and as noise at 32.
   *
   * Two thicker bands still read as the same mark, and the shape survives the downscale.
   */
  const small = Number(size) > 0 && Number(size) <= 48;
  const BANDS = small
    ? [
      { mid: 0.375, amp: 0.055, phase: 0.00, thick: 0.155, alpha: 1.00 },
      { mid: 0.660, amp: 0.055, phase: 0.10, thick: 0.135, alpha: 0.80 },
    ]
    : [
      /* The crest: brightest, and the one that carries the shape. */
      { mid: 0.345, amp: 0.050, phase: 0.00, thick: 0.094, alpha: 1.00 },
      /* The fold behind it. */
      { mid: 0.525, amp: 0.058, phase: 0.09, thick: 0.104, alpha: 0.84 },
      /* And the one falling away into the badge. */
      { mid: 0.706, amp: 0.048, phase: 0.18, thick: 0.082, alpha: 0.58 },
    ];
  /*
   * One wave shared by all three, offset by phase, so the bands run parallel like folds in cloth
   * rather than crossing. A little under one and a quarter cycles across the badge: enough for
   * the curve to read as a wave, not so much that it looks like a squiggle.
   */
  const FREQ = 1.2;
  for (let i = 0; i < BANDS.length; i++) {
    const band = BANDS[i];
    const centre = band.mid + band.amp * Math.sin((bx * FREQ + band.phase) * Math.PI * 2);
    const edge = Math.abs(by - centre) - band.thick / 2;
    // About a pixel of softness at 512px, in badge units.
    const a = clamp01(1 - edge / 0.013);
    if (a > 0) {
      out = over(out, [255, 255, 255, a * band.alpha]);
    }
  }

  return out;
}

/* ------------------------------------------------------------------ *
 * Rasterising
 * ------------------------------------------------------------------ */

function render(size, opts) {
  const inset = (opts && opts.inset) || 0;
  const shadow = !!(opts && opts.shadow);
  // Supersampled, so hard edges inside `sample` still come out antialiased. Large sizes
  // need fewer samples per pixel because each one covers less of the drawing.
  const ss = size >= 512 ? 2 : (size >= 128 ? 6 : 10);
  const px = Buffer.alloc(size * size * 4);

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const c = sample((x + (sx + 0.5) / ss) / size, (y + (sy + 0.5) / ss) / size, inset, shadow, !!(opts && opts.maskable), size, !!(opts && opts.square));
          r += c[0] * c[3];
          g += c[1] * c[3];
          b += c[2] * c[3];
          a += c[3];
        }
      }
      const n = ss * ss;
      const i = (y * size + x) * 4;
      px[i] = a > 0 ? Math.round(r / a) : 0;
      px[i + 1] = a > 0 ? Math.round(g / a) : 0;
      px[i + 2] = a > 0 ? Math.round(b / a) : 0;
      px[i + 3] = Math.round((a / n) * 255);
    }
  }
  return px;
}

/* ------------------------------------------------------------------ *
 * PNG
 * ------------------------------------------------------------------ */

const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) { c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; }
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let crc = -1;
  for (let i = 0; i < buf.length; i++) {
    /*
     * The shift is EIGHT, not one. The table is built with a one-bit shift and the update
     * step is not: writing one here produces a checksum that still looks like a checksum,
     * so every decoder rejects the file with nothing useful to say about why. That is
     * exactly what happened — `file` reported "PNG image data, 180 x 180" while macOS and
     * the image reader both refused to open it.
     */
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buf[i]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

/*
 * CRC-32 checked against the standard test vector before anything is written.
 *
 * This is not decoration: getting it wrong produces files that LOOK valid to every tool
 * that only reads headers, and the failure surfaces as "unsupported or malformed image
 * data" somewhere far away. If the checksum is wrong, nothing here should run at all.
 */
if (crc32(Buffer.from('123456789', 'ascii')) !== 0xcbf43926) {
  throw new Error('CRC-32 self-check failed: got 0x' + crc32(Buffer.from('123456789', 'ascii')).toString(16) + ', wanted 0xcbf43926');
}

function pngChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const body = Buffer.from(type, 'ascii');
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([body, data])), 0);
  return Buffer.concat([len, body, data, crc]);
}

function encodePng(size, rgba) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;    // bit depth
  ihdr[9] = 6;    // truecolour with alpha
  // 10, 11, 12 stay 0: default compression, default filter, no interlace.

  // One filter byte (0 = none) per row. With alpha, "none" is the honest choice.
  const stride = size * 4;
  const raw = Buffer.alloc((stride + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (stride + 1)] = 0;
    rgba.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride);
  }

  return Buffer.concat([
    signature,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

/* ------------------------------------------------------------------ *
 * ICO and ICNS
 * ------------------------------------------------------------------ */

/*
 * ICO. PNG payloads rather than BMP, which is what every browser and Windows itself has
 * accepted since Vista, and which keeps the file a fraction of the size.
 */
function encodeIco(entries) {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);              // reserved
  header.writeUInt16LE(1, 2);              // 1 = icon
  header.writeUInt16LE(entries.length, 4);

  const dir = Buffer.alloc(16 * entries.length);
  let offset = 6 + 16 * entries.length;
  entries.forEach((e, i) => {
    const at = i * 16;
    dir[at] = e.size >= 256 ? 0 : e.size;  // 0 means 256
    dir[at + 1] = e.size >= 256 ? 0 : e.size;
    dir[at + 2] = 0;                        // palette size
    dir[at + 3] = 0;                        // reserved
    dir.writeUInt16LE(1, at + 4);           // colour planes
    dir.writeUInt16LE(32, at + 6);          // bits per pixel
    dir.writeUInt32LE(e.png.length, at + 8);
    dir.writeUInt32LE(offset, at + 12);
    offset += e.png.length;
  });

  return Buffer.concat([header, dir, ...entries.map((e) => e.png)]);
}

/*
 * ICNS. A container of PNG chunks with a four-character type code each. The codes below
 * are Apple's; there is no 1024 point type apart from ic10, which is the 512@2x slot.
 */
function encodeIcns(entries) {
  const chunks = entries.map((e) => {
    const head = Buffer.alloc(8);
    head.write(e.type, 0, 'ascii');
    head.writeUInt32BE(e.png.length + 8, 4);
    return Buffer.concat([head, e.png]);
  });
  const body = Buffer.concat(chunks);
  const head = Buffer.alloc(8);
  head.write('icns', 0, 'ascii');
  head.writeUInt32BE(body.length + 8, 4);
  return Buffer.concat([head, body]);
}

/* ------------------------------------------------------------------ *
 * Write everything
 * ------------------------------------------------------------------ */

const cache = new Map();
function png(size, opts) {
  const key = size + '|' + JSON.stringify(opts || {});
  if (!cache.has(key)) { cache.set(key, encodePng(size, render(size, opts))); }
  return cache.get(key);
}

function write(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, data);
  console.log('  ' + path.relative(APP_DIR, file) + '  ' + (data.length / 1024).toFixed(1) + ' KB');
}

console.log('Drawing the icon...');

// --- what the browser and a phone home screen ask for -------------------------------
const FULL = { inset: 0, shadow: false };
const MASKABLE = { inset: 0, shadow: false, maskable: true };
/*
 * SQUARE: the same drawing, flat and full-bleed, for a phone's home screen.
 *
 * A home screen does NOT want the rounded, transparent-cornered icon that a browser tab does. iOS
 * masks the icon into its own squircle and composites transparency onto BLACK, so a source that is
 * already rounded and transparent gets masked a second time and shows dark artefacts at the corners —
 * or is shrunk inside them. Apple asks for a square, opaque image with no transparency and no
 * pre-rounded corners, and Android wants the same for a legacy icon it will mask itself.
 *
 * The maskable variant is already opaque because Android crops it; this is the same idea for the
 * targets that are NOT maskable but are still a home screen.
 */
const SQUARE = { inset: 0, shadow: false, square: true };

write(path.join(PUBLIC_DIR, 'icon-16.png'), png(16, FULL));
write(path.join(PUBLIC_DIR, 'icon-32.png'), png(32, FULL));
write(path.join(PUBLIC_DIR, 'icon-180.png'), png(180, SQUARE));  // apple-touch-icon
write(path.join(PUBLIC_DIR, 'icon-192.png'), png(192, SQUARE));
write(path.join(PUBLIC_DIR, 'icon-512.png'), png(512, SQUARE));
// Android crops a maskable icon to its own shape, so this one is a full square.
write(path.join(PUBLIC_DIR, 'icon-512-maskable.png'), png(512, MASKABLE));
write(path.join(ASSET_DIR, 'icon-1024.png'), png(1024, FULL));

// The tab icon. Several sizes in one file, which is what makes it look right whether the
// tab is 16px or on a retina display.
write(path.join(PUBLIC_DIR, 'favicon.ico'), encodeIco([16, 32, 48].map((s) => ({ size: s, png: png(s, FULL) }))));

/*
 * Windows. A .bat cannot carry an icon, so the launcher builds a shortcut that points at
 * this file — see "Start NEA Coach.bat". 256 is included because Explorer asks for it.
 */
write(path.join(ASSET_DIR, 'NEA Coach.ico'), encodeIco([16, 32, 48, 64, 128, 256].map((s) => ({ size: s, png: png(s, FULL) }))));

/*
 * macOS. Inset with a shadow, because nothing masks an .icns.
 */
const MAC = { inset: 0.085, shadow: true };
write(path.join(ASSET_DIR, 'NEA Coach.icns'), encodeIcns([
  { type: 'ic11', png: png(32, MAC) },
  { type: 'ic12', png: png(64, MAC) },
  { type: 'ic07', png: png(128, MAC) },
  { type: 'ic08', png: png(256, MAC) },
  { type: 'ic09', png: png(512, MAC) },
  { type: 'ic10', png: png(1024, MAC) },
]));

console.log('Done.');
