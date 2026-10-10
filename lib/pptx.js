'use strict';
/*
 * pptx.js — turn a .pptx into plain structured data.
 *
 * READ-ONLY BY CONSTRUCTION: this module never opens the presentation for
 * writing. It reads the ZIP, parses the XML strings, and returns an object.
 * Nothing it does can alter your file.
 *
 * It extracts, per slide: title, text blocks (with font size and position),
 * images (with their on-slide size), speaker notes, and a content fingerprint.
 * The fingerprint is what lets repeat scans skip unchanged slides and stay
 * cheap.
 */

const path = require('path');
const crypto = require('crypto');
const zipread = require('./zipread');
const imagemeta = require('./imagemeta');

// English Metric Units. PowerPoint stores all geometry in these.
const EMU_PER_INCH = 914400;
// A point is 1/72 inch.
const EMU_PER_POINT = EMU_PER_INCH / 72;

/* ------------------------------------------------------------------ *
 * Namespace-agnostic XML helpers.
 * We cannot assume which prefixes a given producer used, so every helper
 * tolerates any prefix or none.
 * ------------------------------------------------------------------ */

/** Decode the five XML entities that appear in text runs. */
function decodeXml(s) {
  if (s == null) return '';
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&amp;/g, '&'); // must be last so we do not double-decode
}

/** Return every match of <tag ...> ... </tag>, non-greedy, with its inner XML. */
function elements(xml, localName) {
  const out = [];
  const re = new RegExp(`<(?:[A-Za-z0-9_.-]+:)?${localName}(?:\\s[^>]*)?>([\\s\\S]*?)</(?:[A-Za-z0-9_.-]+:)?${localName}>`, 'g');
  let m;
  while ((m = re.exec(xml)) !== null) out.push(m[1]);
  return out;
}

/** Return the first matching element's inner XML, or null. */
function firstElement(xml, localName) {
  const all = elements(xml, localName);
  return all.length ? all[0] : null;
}

/** Collect the full opening tags (attributes included) for a local name. */
function openTags(xml, localName) {
  const out = [];
  const re = new RegExp(`<(?:[A-Za-z0-9_.-]+:)?${localName}(\\s[^>]*?)?/?>`, 'g');
  let m;
  while ((m = re.exec(xml)) !== null) out.push(m[1] || '');
  return out;
}

/** Read one attribute out of an attribute string. */
function attr(attrs, name) {
  if (!attrs) return null;
  const re = new RegExp(`${name}\\s*=\\s*"([^"]*)"`);
  const m = re.exec(attrs);
  return m ? m[1] : null;
}

/**
 * Read the attribute string of the FIRST occurrence of a tag, whether it is
 * self-closing or paired, searching only from the start of the given XML.
 *
 * This exists because `<a:rPr .../>` is normally self-closing in real PowerPoint
 * files, and the `elements()` helper only matches paired tags. Missing it loses
 * font sizes, bold flags and run language.
 */
function readSelfClosingOrOpen(xml, localName) {
  if (!xml) return '';
  const re = new RegExp(`^\\s*<(?:[A-Za-z0-9_.-]+:)?${localName}(\\s[^>]*?)?/?>`, 'i');
  const m = re.exec(xml);
  return m && m[1] ? m[1] : '';
}

/** Read a numeric attribute. */
function numAttr(attrs, name) {
  const v = attr(attrs, name);
  if (v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

/**
 * Text inside <a:t> runs. The negative lookahead keeps us from matching
 * <a:tbl>, <a:tc> and friends, which are tables, not text runs.
 */
function extractRuns(xml) {
  const out = [];
  const re = /<(?:[A-Za-z0-9_.-]+:)?t(\s[^>]*)?>([\s\S]*?)<\/(?:[A-Za-z0-9_.-]+:)?t>/g;
  let m;
  while ((m = re.exec(xml)) !== null) out.push(decodeXml(m[2]));
  return out;
}

/* ------------------------------------------------------------------ *
 * Text extraction
 * ------------------------------------------------------------------ */

const BULLET_HINT = /^\s*(?:[-\u2022\u25cf\u25aa\u2023\u2043*>]|\d+[.)])\s+/;

/**
 * Split an XML fragment into its direct children with a given local name.
 *
 * `elements()` also walks nested descendants, which is wrong for lists like
 * <a:tr> inside <a:tbl>: it would flatten the whole grid into one list. This
 * walks one level only, and counts self-closing tags as children so that cell
 * positions stay aligned.
 */
function childElements(xml, localName) {
  const out = [];
  const re = new RegExp(`<(?:[A-Za-z0-9_.-]+:)?${localName}(\\s[^>]*?)?(\\/?)>`, 'g');
  let m;
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1] || '';
    if (m[2] === '/') {
      out.push({ xml: '', attrs });
      continue;
    }
    const closeRe = new RegExp(`</(?:[A-Za-z0-9_.-]+:)?${localName}>`);
    const rest = xml.slice(re.lastIndex);
    const closeM = closeRe.exec(rest);
    out.push({ xml: closeM ? rest.slice(0, closeM.index) : rest, attrs });
    if (closeM) re.lastIndex += closeM.index + closeM[0].length;
  }
  return out;
}

/**
 * Parse an <a:tbl> into rows of cells.
 *
 * Tables used to be detected but never read: the parser recorded "this page has
 * a table" and threw the contents away. On the real 60-page portfolio that made
 * 9 pages — including four complete manufacture plans — look empty, and the AI
 * reviewing them reported them as blank pages that needed writing. Table text is
 * now extracted as ordinary paragraphs so it flows into the word count, the
 * local checks and the AI prompt exactly like any other text.
 */
function extractTable(shapeXml) {
  const tblXml = firstElement(shapeXml, 'tbl');
  if (tblXml == null) return null;

  const rows = [];
  for (const tr of childElements(tblXml, 'tr')) {
    const cells = [];
    for (const tc of childElements(tr.xml, 'tc')) {
      /*
       * <a:t> runs give the text; <a:br/> is an explicit line break inside a
       * cell and must become a space or words would run together.
       */
      const text = decodeXml(
        tc.xml.replace(/<(?:[A-Za-z0-9_.-]+:)?br\b[^>]*\/?>/g, ' ').replace(/<[^>]+>/g, ' ')
      )
        .replace(/\s+/g, ' ')
        .trim();
      const szMatch = /<(?:[A-Za-z0-9_.-]+:)?rPr\b[^>]*\bsz="(\d+)"/.exec(tc.xml);
      const sizePt = szMatch ? Number(szMatch[1]) / 100 : null;
      const gridSpan = numAttr(tc.attrs, 'gridSpan');
      const rowSpan = numAttr(tc.attrs, 'rowSpan');
      cells.push({
        text,
        sizePt: Number.isFinite(sizePt) && sizePt > 0 ? sizePt : null,
        // Merged cells repeat text in the archive; hMerge marks the continuation.
        hMerge: /\bhMerge="1"/.test(tc.attrs),
        gridSpan: gridSpan && gridSpan > 1 ? gridSpan : 1,
        rowSpan: rowSpan && rowSpan > 1 ? rowSpan : 1,
      });
    }
    rows.push(cells);
  }

  // Grid columns carry real widths in EMU, which lets the AI judge proportion.
  const grid = firstElement(tblXml, 'tblGrid');
  const colWidths = grid
    ? openTags(grid, 'gridCol')
        .map((a) => numAttr(a, 'w'))
        .filter((w) => w != null)
    : [];

  return { rows, colWidths };
}

/**
 * Render a parsed table as aligned plain text. Prose beats markup here: this is
 * what the AI actually reads, so rows become lines and cells become " | ".
 */
function tableToText(table) {
  if (!table || !table.rows.length) return '';
  return table.rows.map((cells) => cells.map((c) => c.text).join(' | ')).join('\n');
}

/**
 * Pull paragraphs out of a shape's text body.
 * Font size resolution follows PowerPoint's own inheritance: run properties
 * win, then paragraph properties, then the shape default.
 */
/**
 * Parse a shape's text body.
 *
 * @param {string} shapeXml
 * @param {object} [inherited] Placeholder defaults from the slide layout/master,
 *   used when the slide itself does not state a size. PowerPoint omits these
 *   constantly, so without it the size of most body text is simply unknown.
 */
function extractTextBody(shapeXml, inherited) {
  const txBody = firstElement(shapeXml, 'txBody');
  if (txBody == null) return null;
  const inherit = inherited || {};

  /*
   * Autofit matters when judging overflow. A box with <a:spAutoFit/> GROWS to fit
   * its text, so text extending past the recorded height is normal and NOT a
   * defect. <a:normAutofit/> instead SHRINKS the text to fit, which can make it
   * unreadably small (fontScale). A fixed box with neither genuinely clips.
   * This distinction was verified against a real render before being trusted.
   */
  // These live inside <a:bodyPr> as self-closing tags, so we search the whole
  // shape rather than only the extracted text body.
  let autofit = 'none';
  if (/<(?:[A-Za-z0-9_.-]+:)?spAutoFit\b/.test(shapeXml)) autofit = 'shape-grows';
  else if (/<(?:[A-Za-z0-9_.-]+:)?normAutofit\b/.test(shapeXml)) autofit = 'text-shrinks';
  const normAutofitMatch = /<(?:[A-Za-z0-9_.-]+:)?normAutofit\b([^>]*?)\/?>/.exec(shapeXml);
  const fontScale = normAutofitMatch ? numAttr(normAutofitMatch[1] || '', 'fontScale') : null;
  const lnSpcReduction = normAutofitMatch ? numAttr(normAutofitMatch[1] || '', 'lnSpcReduction') : null;

  const defRunProps = firstElement(txBody, 'defRPr');
  let defaultSize = defRunProps ? numAttr(openTags(defRunProps, 'rPr')[0] || '', 'sz') : null;

  /*
   * Size resolution order, following PowerPoint's own precedence:
   *   run rPr -> paragraph defRPr -> txBody lstStyle level -> layout/master.
   * The layout/master fallback is the one that matters in practice, because
   * PowerPoint usually omits sizes in the slide's own XML entirely.
   */
  const ownLevelSizes = readLevelSizes(firstElement(txBody, 'lstStyle') || '');
  const levelSizes = Object.assign({}, inherit.levelSizes || {}, ownLevelSizes);
  if (defaultSize == null && inherit.sizePt != null) defaultSize = inherit.sizePt * 100;

  const sizeFor = (runSize, paraSize, lvl) => {
    const hundredths = runSize != null ? runSize
      : paraSize != null ? paraSize
      : levelSizes[lvl] != null ? levelSizes[lvl] * 100
      : defaultSize;
    return hundredths != null ? hundredths / 100 : null;
  };

  const paragraphs = [];
  for (const pXml of elements(txBody, 'p')) {
    const pPr = firstElement(pXml, 'pPr');
    const pSize = pPr ? numAttr(openTags(pPr, 'defRPr')[0] || '', 'sz') : null;
    const pAlign = pPr ? attr(openTags(pPr, 'pPr')[0] || '', 'algn') : null;
    const pLevel = pPr ? (numAttr(openTags(pPr, 'pPr')[0] || '', 'lvl') || 0) : 0;

    const runs = [];
    for (const rXml of elements(pXml, 'r')) {
      /*
       * Read rPr positionally, from the start of the run.
       *
       * PowerPoint writes run properties as a SELF-CLOSING tag, e.g.
       *   <a:r><a:rPr lang="en-GB" sz="1200" b="1"/><a:t>Collecting data:</a:t></a:r>
       * Our <tag>...</tag> matcher cannot see self-closing elements, so a naive
       * lookup returns null, which loses both the font size and the text itself.
       * On a real portfolio that silently disabled the text-size check entirely.
       */
      const rPrAttrs = readSelfClosingOrOpen(rXml, 'rPr');
      const text = extractRuns(rXml).join('');
      if (!text) continue;

      const size = numAttr(rPrAttrs, 'sz');
      const bold = attr(rPrAttrs, 'b');
      const lang = attr(rPrAttrs, 'lang');

      runs.push({
        text,
        sizePt: sizeFor(size, pSize, pLevel),
        bold: bold === '1' || bold === 'true',
        lang: lang || null,
      });
    }

    // A paragraph can also carry text directly through fields (slide numbers).
    if (!runs.length) {
      const fldText = extractRuns(pXml).join('');
      if (fldText) {
        runs.push({
          text: fldText,
          sizePt: sizeFor(null, pSize, pLevel),
          bold: false,
          lang: null,
        });
      }
    }

    if (!runs.length) {
      paragraphs.push({ text: '', runs: [], sizePt: sizeFor(null, pSize, pLevel), align: pAlign, bulleted: false, level: pLevel });
      continue;
    }

    const text = runs.map((r) => r.text).join('');
    const lvl = pPr ? numAttr(openTags(pPr, 'pPr')[0] || '', 'lvl') : null;
    const bulleted = BULLET_HINT.test(text) || (pPr ? /buChar|buAutoNum/.test(pPr) : false);

    paragraphs.push({
      text,
      runs,
      sizePt: sizeFor(null, pSize, lvl == null ? pLevel : lvl),
      align: pAlign,
      bulleted,
      level: lvl == null ? 0 : lvl,
    });
  }

  if (!paragraphs.length) return null;

  const sizes = {};
  const fonts = {};
  let charCount = 0;
  for (const p of paragraphs) {
    charCount += p.text.length;
    for (const r of p.runs) {
      if (r.sizePt != null) sizes[r.sizePt] = (sizes[r.sizePt] || 0) + r.text.length;
    }
  }

  return {
    paragraphs,
    sizes, // sizePt -> characters at that size
    charCount,
    wordCount: paragraphs.map((p) => p.text).join(' ').trim().split(/\s+/).filter(Boolean).length,
    fonts,
    // Passed through so the overflow rule can avoid false alarms on boxes that
    // are allowed to grow, and instead flag text shrunk to nothing.
    autofit,
    fontScale,
    lnSpcReduction,
  };
}

/* ------------------------------------------------------------------ *
 * Geometry
 * ------------------------------------------------------------------ */

/**
 * Convert a shape's transform into inches and points.
 *
 * IMPORTANT: read <a:off>/<a:ext> only from inside <a:xfrm>. Searching the whole
 * shape is wrong: <p:cNvPr> commonly carries an <a:extLst> extension block, and a
 * loose match on "ext" happily captures <a:extLst> instead of <a:ext>. That bug
 * made every shape on a real 60-page portfolio report y=0, h=0, which silently
 * disabled the off-canvas and overflow checks.
 */
function extractGeometry(shapeXml) {
  const xfrm = firstElement(shapeXml, 'xfrm');
  if (xfrm == null) return null;

  const offAttrs = openTags(xfrm, 'off')[0];
  const extAttrs = openTags(xfrm, 'ext')[0];
  if (!offAttrs || !extAttrs) return null;

  const x = numAttr(offAttrs, 'x');
  const y = numAttr(offAttrs, 'y');
  const cx = numAttr(extAttrs, 'cx');
  const cy = numAttr(extAttrs, 'cy');
  if (cx == null || cy == null) return null;

  return {
    xIn: x != null ? x / EMU_PER_INCH : null,
    yIn: y != null ? y / EMU_PER_INCH : null,
    wIn: cx / EMU_PER_INCH,
    hIn: cy / EMU_PER_INCH,
    xPt: x != null ? x / EMU_PER_POINT : null,
    yPt: y != null ? y / EMU_PER_POINT : null,
    wPt: cx / EMU_PER_POINT,
    hPt: cy / EMU_PER_POINT,
  };
}

/* ------------------------------------------------------------------ *
 * Relationships
 * ------------------------------------------------------------------ */

function parseRels(xml) {
  const map = new Map();
  if (!xml) return map;
  const re = /<Relationship\b([^>]*)>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const a = m[1];
    const id = attr(a, 'Id');
    if (!id) continue;
    map.set(id, {
      id,
      type: attr(a, 'Type') || '',
      target: attr(a, 'Target') || '',
      mode: attr(a, 'TargetMode') || 'Internal',
    });
  }
  return map;
}

/** Resolve a relationship target to a path inside the archive. */
function resolveTarget(baseDir, target) {
  if (/^[a-z]+:\/\//i.test(target)) return null; // external link, not in the archive
  if (target.startsWith('/')) return target.slice(1);
  const joined = path.posix.join(baseDir, target);
  return joined;
}

/* ------------------------------------------------------------------ *
 * Layout and master inheritance
 *
 * A real PowerPoint slide often stores almost nothing: the title run can be
 * literally <a:rPr lang="en-GB"/> with no size, because the size lives in the
 * slide layout, and the layout inherits from the master. Without this, every
 * inherited text size reads as unknown and every inherited placeholder has no
 * position, which silently disables the text-size, off-canvas and overflow
 * checks on exactly the decks that matter.
 * ------------------------------------------------------------------ */

/** Read the <a:lvlNpPr> default sizes out of a style block, keyed by level. */
function readLevelSizes(styleXml) {
  const out = {};
  if (!styleXml) return out;
  for (let lvl = 1; lvl <= 9; lvl++) {
    const m = new RegExp(`<(?:[A-Za-z0-9_.-]+:)?lvl${lvl}pPr([^>]*)>`).exec(styleXml);
    if (!m) continue;
    // The size can sit on lvlNpPr itself or on its <a:defRPr> child.
    let sz = numAttr(m[1], 'sz');
    if (sz == null) {
      const rest = styleXml.slice(m.index + m[0].length);
      const dr = openTags(rest.slice(0, 400), 'defRPr')[0];
      if (dr) sz = numAttr(dr, 'sz');
    }
    if (sz != null) out[lvl - 1] = sz / 100;
  }
  return out;
}

/** Pull placeholder defaults (position, size, autofit, level sizes) from a layout/master part. */
function readPlaceholderStyles(xml) {
  const byKey = {};
  if (!xml) return byKey;

  const shapeRe = /<(?:[A-Za-z0-9_.-]+:)?sp\b[\s>][\s\S]*?<\/(?:[A-Za-z0-9_.-]+:)?sp>/g;
  let m;
  while ((m = shapeRe.exec(xml)) !== null) {
    const sp = m[0];
    const phAttrs = openTags(sp, 'ph')[0];
    if (!phAttrs) continue;

    const type = (attr(phAttrs, 'type') || 'body').toLowerCase();
    const idx = attr(phAttrs, 'idx') || '0';
    const key = `${type}:${idx}`;

    const txBody = firstElement(sp, 'txBody') || '';
    const lstStyle = firstElement(txBody, 'lstStyle') || '';
    const pPr = firstElement(txBody, 'pPr') || '';

    // Size can live on <a:lvl1pPr>, or on a paragraph's <a:defRPr>.
    let sizePt = null;
    const lvl1 = openTags(lstStyle, 'lvl1pPr')[0];
    if (lvl1) sizePt = numAttr(lvl1, 'sz') != null ? numAttr(lvl1, 'sz') / 100 : null;
    if (sizePt == null && pPr) {
      const dr = openTags(pPr, 'defRPr')[0];
      if (dr && numAttr(dr, 'sz') != null) sizePt = numAttr(dr, 'sz') / 100;
    }

    byKey[key] = {
      geometry: extractGeometry(sp),
      sizePt,
      levelSizes: readLevelSizes(lstStyle),
      autofit: /<(?:[A-Za-z0-9_.-]+:)?spAutoFit\b/.test(sp)
        ? 'shape-grows'
        : /<(?:[A-Za-z0-9_.-]+:)?normAutofit\b/.test(sp)
        ? 'text-shrinks'
        : /<(?:[A-Za-z0-9_.-]+:)?noAutofit\b/.test(sp)
        ? 'none'
        : null,
      name: (openTags(sp, 'cNvPr')[0] && attr(openTags(sp, 'cNvPr')[0], 'name')) || null,
    };
  }
  return byKey;
}

/** Build layout and master placeholder lookups for the whole presentation. */
function buildLayoutResolver(zip) {
  const layouts = new Map(); // layout path -> {byKey, masterPath}
  const masters = new Map(); // master path -> {byKey}

  // Masters first.
  for (const name of zip.list()) {
    if (!/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name)) continue;
    const xml = zip.readText(name);
    if (!xml) continue;
    const base = { byKey: readPlaceholderStyles(xml) };
    const txStyles = firstElement(xml, 'txStyles');
    if (txStyles) {
      base.titleSizes = readLevelSizes(firstElement(txStyles, 'titleStyle') || '');
      base.bodySizes = readLevelSizes(firstElement(txStyles, 'bodyStyle') || '');
    }
    masters.set(name, base);
  }

  // Then layouts, each pointing at its master.
  for (const name of zip.list()) {
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)) continue;
    const xml = zip.readText(name);
    if (!xml) continue;
    const dir = path.posix.dirname(name);
    const base = path.posix.basename(name);
    const rels = parseRels(zip.readText(`${dir}/_rels/${base}.rels`));
    let masterPath = null;
    for (const r of rels.values()) {
      if (/slideMaster/.test(r.type)) {
        masterPath = resolveTarget(dir, r.target);
        break;
      }
    }
    layouts.set(name, { byKey: readPlaceholderStyles(xml), masterPath });
  }

  return { layouts, masters };
}

/**
 * Resolve what a slide's placeholder should inherit.
 * Layout wins over master, matching PowerPoint's own precedence.
 */
function resolvePlaceholder(resolver, layoutPath, type, idx) {
  const key = `${String(type).toLowerCase()}:${idx == null ? '0' : idx}`;
  const blank = { geometry: null, sizePt: null, levelSizes: {}, autofit: null, name: null };

  const layout = layoutPath ? resolver.layouts.get(layoutPath) : null;
  const fromLayout = layout && layout.byKey[key] ? layout.byKey[key] : null;

  const masterPath = layout ? layout.masterPath : null;
  const master = masterPath ? resolver.masters.get(masterPath) : null;
  const fromMaster = master && master.byKey[key] ? master.byKey[key] : null;

  const merge = (a, b) => ({
    geometry: (a && a.geometry) || (b && b.geometry) || null,
    sizePt: a && a.sizePt != null ? a.sizePt : b && b.sizePt != null ? b.sizePt : null,
    levelSizes: Object.assign({}, (b && b.levelSizes) || {}, (a && a.levelSizes) || {}),
    autofit: (a && a.autofit) || (b && b.autofit) || null,
    name: (a && a.name) || (b && b.name) || null,
  });

  const out = merge(fromLayout, fromMaster);

  // Fall back to the master's overall title/body text styles.
  if (out.sizePt == null && master) {
    const styleSet = /title|ctrTitle/.test(key) ? master.titleSizes : master.bodySizes;
    if (styleSet && styleSet[0] != null) out.sizePt = styleSet[0];
    if (styleSet) out.levelSizes = Object.assign({}, styleSet, out.levelSizes);
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * Main parse
 * ------------------------------------------------------------------ */

function parsePresentation(filePath) {
  const zip = zipread.open(filePath);

  // --- Slide size comes from the presentation part.
  const presXml = zip.readText('ppt/presentation.xml');
  if (!presXml) throw new Error('Not a PowerPoint file: ppt/presentation.xml is missing');

  const sldSzAttrs = openTags(presXml, 'sldSz')[0];
  const cxEmu = numAttr(sldSzAttrs || '', 'cx') || 9144000;
  const cyEmu = numAttr(sldSzAttrs || '', 'cy') || 6858000;
  const slideW = cxEmu / EMU_PER_INCH;
  const slideH = cyEmu / EMU_PER_INCH;
  const slideWPt = cxEmu / EMU_PER_POINT;
  const slideHPt = cyEmu / EMU_PER_POINT;

  const presRels = parseRels(zip.readText('ppt/_rels/presentation.xml.rels'));

  // --- Display order. This is NOT filename order: it follows the sldId
  // Layout and master placeholder defaults, read once and reused per slide.
  const resolver = buildLayoutResolver(zip);

  //     list plus each sldId's r:id. Order matters, because "slide 7" must
  //     mean the same page the student sees.
  const sldIdLst = firstElement(presXml, 'sldIdLst');
  const ordered = [];
  if (sldIdLst) {
    const re = /<(?:[A-Za-z0-9_.-]+:)?sldId\b([^>]*)>/g;
    let m;
    while ((m = re.exec(sldIdLst)) !== null) {
      const rid = attr(m[1], 'r:id') || attr(m[1], 'id');
      const rel = presRels.get(rid);
      if (rel) {
        const p = resolveTarget('ppt', rel.target);
        if (p) ordered.push(p);
      }
    }
  }

  // Fall back to numeric filename order if the relationship walk found nothing.
  if (!ordered.length) {
    for (const name of zip.list()) {
      if (/^ppt\/slides\/slide\d+\.xml$/.test(name)) ordered.push(name);
    }
    ordered.sort((a, b) => {
      const na = Number((a.match(/slide(\d+)\.xml$/) || [])[1] || 0);
      const nb = Number((b.match(/slide(\d+)\.xml$/) || [])[1] || 0);
      return na - nb;
    });
  }

  const slides = [];
  const mediaCache = new Map(); // archive path -> dimensions or null

  for (let i = 0; i < ordered.length; i++) {
    const slidePath = ordered[i];
    const slideXml = zip.readText(slidePath);
    if (!slideXml) continue;

    const dir = path.posix.dirname(slidePath);
    const base = path.posix.basename(slidePath);
    const relsPath = `${dir}/_rels/${base}.rels`;
    const rels = parseRels(zip.readText(relsPath));

    // Which layout does this slide use? Needed for inherited sizes and positions.
    let layoutPath = null;
    for (const r of rels.values()) {
      if (/slideLayout/.test(r.type)) {
        layoutPath = resolveTarget(dir, r.target);
        break;
      }
    }

    const shapes = [];
    const images = [];
    const hashParts = [slideXml];
    let hasTitlePlaceholder = false;
    let titleText = null;

    // A top-level <p:sp> or <p:pic> is one shape. We match them individually
    // rather than splitting on <p:spTree>, which keeps group shapes intact.
    const shapeRe = /<(?:[A-Za-z0-9_.-]+:)?(sp|pic|graphicFrame|grpSp)\b[\s>][\s\S]*?<\/(?:[A-Za-z0-9_.-]+:)?\1>/g;
    let sm;
    while ((sm = shapeRe.exec(slideXml)) !== null) {
      const kind = sm[1];
      const shapeXml = sm[0];
      const ownGeom = extractGeometry(shapeXml);

      const nvAttrs = openTags(shapeXml, 'cNvPr')[0] || '';
      const shapeName = attr(nvAttrs, 'name') || '';
      const shapeId = numAttr(nvAttrs, 'id');

      // Placeholder type tells us whether this shape is the title.
      const phAttrs = openTags(shapeXml, 'ph')[0];
      const phTypeRaw = phAttrs ? attr(phAttrs, 'type') : null;
      const phIdx = phAttrs ? attr(phAttrs, 'idx') : null;
      const isTitlePh = phTypeRaw === 'title' || phTypeRaw === 'ctrTitle' || /^title\b/i.test(shapeName);

      /*
       * Inherit from the layout/master. This is where PowerPoint keeps the real
       * font sizes and often the exact placeholder position.
       */
      const inherited = phAttrs
        ? resolvePlaceholder(resolver, layoutPath, phTypeRaw || 'body', phIdx)
        : { geometry: null, sizePt: null, levelSizes: {}, autofit: null };

      // Layout geometry fills in only when the slide states none of its own.
      const geom = ownGeom || inherited.geometry;

      if (kind === 'pic') {
        const blipAttrs = openTags(shapeXml, 'blip')[0] || '';
        const embed = attr(blipAttrs, 'r:embed') || attr(blipAttrs, 'embed');
        const rel = embed ? rels.get(embed) : null;
        const target = rel ? resolveTarget(dir, rel.target) : null;
        let dims = null;

        if (target && zip.has(target)) {
          if (!mediaCache.has(target)) {
            /*
             * Read only the header. The largest image format header we parse is
             * a 44-byte EMF/WMF record; 4 KB is ample for PNG, JPEG, GIF and BMP
             * dimensions. On a deck with 500 MB of images this avoids inflating
             * media we only need a few bytes of.
             */
            const bytes = zip.read(target, 4096);
            const vector = imagemeta.getVectorInfo(bytes);
            mediaCache.set(target, vector || imagemeta.getDimensions(bytes));
          }
          dims = mediaCache.get(target);
        }

        // Effective DPI = pixels across / inches across on the slide.
        let effectiveDpi = null;
        if (dims && dims.width && geom && geom.wIn > 0.01) {
          effectiveDpi = Math.round(dims.width / geom.wIn);
        }

        images.push({
          shapeName,
          shapeId,
          archivePath: target,
          pixelWidth: dims ? dims.width : null,
          pixelHeight: dims ? dims.height : null,
          format: dims ? dims.type : null,
          isVector: !!(dims && (dims.type === 'emf' || dims.type === 'wmf')),
          placedWidthIn: geom ? Number(geom.wIn.toFixed(2)) : null,
          placedHeightIn: geom ? Number(geom.hIn.toFixed(2)) : null,
          effectiveDpi,
          xIn: geom ? Number(geom.xIn.toFixed(2)) : null,
          yIn: geom ? Number(geom.yIn.toFixed(2)) : null,
          /*
           * Alt text is the only description of a picture's CONTENT available
           * without rendering the page. Students rarely write it, but when they
           * do it is the difference between the AI knowing a page shows a
           * dimensioned working drawing and knowing only that a JPEG exists.
           */
          altText: (() => {
            const nv = openTags(shapeXml, 'cNvPr')[0];
            if (!nv) return null;
            const d = attr(nv, 'descr');
            return d && d.trim() ? decodeXml(d).trim() : null;
          })(),
        });

        if (target) hashParts.push(target);
        shapes.push({ kind, name: shapeName, id: shapeId, isTitle: false, geometry: geom, hasText: false });
        continue;
      }

      if (kind === 'graphicFrame') {
        // Tables and charts live here.
        const hasChart = /graphicData[^>]*uri="[^"]*chart/.test(shapeXml);
        const table = hasChart ? null : extractTable(shapeXml);

        if (table && table.rows.length) {
          /*
           * Table text becomes ordinary paragraphs. Row and cell sizes default to
           * null rather than a guessed point size: a wrong size here would create
           * false "text too small" findings on every table in the deck.
           */
          const paragraphs = [];
          for (const cells of table.rows) {
            for (const cell of cells) {
              if (!cell.text) continue;
              if (cell.hMerge) continue; // merged continuation, not new content
              paragraphs.push({
                text: cell.text,
                runs: [],
                sizePt: cell.sizePt,
                align: null,
                bulleted: false,
                level: 0,
                isTableCell: true,
              });
            }
          }
          if (paragraphs.length) {
            const text = paragraphs.map((p) => p.text).join('\n');
            shapes.push({
              kind: 'table',
              name: shapeName,
              id: shapeId,
              isTitle: false,
              geometry: geom,
              hasText: true,
              paragraphs,
              charCount: text.length,
              wordCount: text.split(/\s+/).filter(Boolean).length,
              table: {
                rowCount: table.rows.length,
                colCount: table.colWidths.length || (table.rows[0] ? table.rows[0].length : 0),
                colWidths: table.colWidths,
              },
              autofit: 'none',
              fontScale: null,
            });
            continue;
          }
        }

        shapes.push({
          kind: hasChart ? 'chart' : table ? 'table-empty' : 'graphic',
          name: shapeName,
          id: shapeId,
          isTitle: false,
          geometry: geom,
          hasText: false,
        });
        continue;
      }

      // Plain shape: may or may not carry text.
      const body = extractTextBody(shapeXml, inherited);
      const noteRels = openTags(shapeXml, 'hlinkClick')[0];
      const linkRel = noteRels ? rels.get(attr(noteRels, 'r:id') || '') : null;

      if (body && body.charCount > 0) {
        const plainText = body.paragraphs.map((p) => p.text).join('\n').trim();
        if (isTitlePh && !titleText) {
          titleText = plainText;
          hasTitlePlaceholder = true;
        }
        shapes.push({
          kind: 'text',
          name: shapeName,
          id: shapeId,
          isTitle: isTitlePh,
          placeholderType: phTypeRaw,
          geometry: geom,
          hasText: true,
          paragraphs: body.paragraphs,
          sizes: body.sizes,
          charCount: body.charCount,
          wordCount: body.wordCount,
          /*
           * Autofit can be inherited too: a placeholder with no <a:bodyPr> of its
           * own takes the layout's setting. Treating "unknown" as "fixed box"
           * would produce false overflow warnings, so fall back explicitly.
           */
          autofit: body.autofit && body.autofit !== 'none' ? body.autofit : (inherited.autofit || 'none'),
          fontScale: body.fontScale,
          lnSpcReduction: body.lnSpcReduction,
          hyperlink: linkRel ? linkRel.target : null,
        });
      } else {
        shapes.push({
          kind: 'shape',
          name: shapeName,
          id: shapeId,
          isTitle: isTitlePh,
          placeholderType: phTypeRaw,
          geometry: geom,
          hasText: false,
        });
      }
    }

    // Speaker notes.
    const notesPath = `${dir}/notesSlides/notesSlide${(slidePath.match(/slide(\d+)\.xml$/) || [])[1]}.xml`;
    let notes = '';
    const notesRel = Array.from(rels.values()).find((r) => /notesSlide/.test(r.type));
    const actualNotesPath = notesRel ? resolveTarget(dir, notesRel.target) : notesPath;
    if (actualNotesPath && zip.has(actualNotesPath)) {
      const notesXml = zip.readText(actualNotesPath) || '';
      notes = extractRuns(notesXml).join(' ').replace(/\s+/g, ' ').trim();
    }

    // Aggregate the slide's text.
    const textShapes = shapes.filter((s) => s.hasText);
    const bodyShapes = textShapes.filter((s) => !s.isTitle);
    const allText = textShapes
      .map((s) => s.paragraphs.map((p) => p.text).join('\n'))
      .join('\n')
      .trim();

    // If there was no real title placeholder, fall back to the topmost short
    // text block, which is what a human would read as the title.
    if (!titleText) {
      const candidates = bodyShapes
        .filter((s) => s.geometry && s.wordCount <= 14 && s.charCount <= 90)
        .sort((a, b) => (a.geometry.yIn ?? 99) - (b.geometry.yIn ?? 99));
      if (candidates.length && candidates[0].geometry && candidates[0].geometry.yIn < slideH * 0.35) {
        titleText = candidates[0].paragraphs.map((p) => p.text).join(' ').trim();
      }
    }

    const allSizes = {};
    const allSizeParaCounts = {};
    const fontFaces = {};
    for (const s of textShapes) {
      /*
       * Build the size histogram from the resolved paragraph sizes. An earlier
       * version read s.sizes (run-level counts), but inherited sizes arrive at
       * paragraph level, so every inherited size was missing from the histogram
       * and the "text too small" summary had nothing to report.
       */
      for (const p of s.paragraphs) {
        const size = p.sizePt != null ? p.sizePt : (p.runs[0] && p.runs[0].sizePt);
        if (size != null) {
          const chars = Math.max(1, (p.text || '').length);
          allSizes[size] = (allSizes[size] || 0) + chars;
          allSizeParaCounts[size] = (allSizeParaCounts[size] || 0) + 1;
        }
        for (const r of p.runs) {
          if (r.lang) fontFaces[r.lang] = (fontFaces[r.lang] || 0) + 1;
        }
      }
      if (s.sizes) {
        for (const [sz, count] of Object.entries(s.sizes)) {
          if (allSizes[sz] == null) allSizes[sz] = count;
        }
      }
    }

    const wordCount = allText ? allText.split(/\s+/).filter(Boolean).length : 0;
    const longestParagraphWords = textShapes.reduce((max, s) => {
      for (const p of s.paragraphs) {
        const w = p.text.split(/\s+/).filter(Boolean).length;
        if (w > max) max = w;
      }
      return max;
    }, 0);

    // Fingerprint: the slide XML plus the bytes of every image it uses. If the
    // student changes a picture without touching the text, this still changes,
    // so the AI gets asked about it again.
    const fingerprint = crypto
      .createHash('sha1')
      .update(hashParts.join('|'))
      .digest('hex')
      .slice(0, 16);

    slides.push({
      index: i + 1,
      archivePath: slidePath,
      title: titleText || null,
      hasTitle: !!titleText,
      hasTitlePlaceholder,
      text: allText,
      textShapes: textShapes.map((s) => ({
        kind: s.kind,
        name: s.name,
        isTitle: s.isTitle,
        geometry: s.geometry,
        paragraphs: s.paragraphs.map((p) => ({
          text: p.text,
          sizePt: p.sizePt != null ? p.sizePt : (p.runs[0] && p.runs[0].sizePt) || null,
          bulleted: p.bulleted,
        })),
        charCount: s.charCount,
        wordCount: s.wordCount,
        table: s.table || null,
        // 'shape-grows' means overflow is not a defect; see extractTextBody.
        autofit: s.autofit || 'none',
        fontScale: s.fontScale != null ? s.fontScale : null,
      })),
      images,
      shapes: shapes.map((s) => ({
        kind: s.kind,
        name: s.name,
        isTitle: !!s.isTitle,
        geometry: s.geometry,
        hasText: !!s.hasText,
      })),
      notes,
      notesWordCount: notes ? notes.split(/\s+/).filter(Boolean).length : 0,
      wordCount,
      charCount: allText.length,
      longestParagraphWords,
      sizesPt: Object.keys(allSizes).map(Number).sort((a, b) => a - b),
      sizeHistogram: allSizes,
      sizeParagraphCounts: allSizeParaCounts,
      langs: fontFaces,
      imageCount: images.length,
      shapeCount: shapes.length,
      fingerprint,
    });
  }

  const coreXml = zip.readText('docProps/core.xml');
  const appXml = zip.readText('docProps/app.xml');
  const readProp = (xml, tag) => {
    if (!xml) return null;
    const m = new RegExp(`<(?:[A-Za-z0-9_.-]+:)?${tag}[^>]*>([\\s\\S]*?)</(?:[A-Za-z0-9_.-]+:)?${tag}>`).exec(xml);
    return m ? decodeXml(m[1]).trim() : null;
  };

  return {
    filePath,
    fileName: path.basename(filePath),
    slideWidthIn: slideW,
    slideHeightIn: slideH,
    slideWidthPt: slideWPt,
    slideHeightPt: slideHPt,
    aspect: slideW / slideH,
    slides,
    slideCount: slides.length,
    meta: {
      title: readProp(coreXml, 'title'),
      author: readProp(coreXml, 'creator'),
      lastModifiedBy: readProp(coreXml, 'lastModifiedBy'),
      modified: readProp(coreXml, 'modified'),
      revision: readProp(coreXml, 'revision'),
      app: readProp(appXml, 'Application'),
      company: readProp(appXml, 'Company'),
    },
  };
}

module.exports = { parsePresentation, EMU_PER_INCH, EMU_PER_POINT };
