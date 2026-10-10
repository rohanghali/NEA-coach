'use strict';

/*
 * Turn a PDF into something the free structural checks can read.
 *
 * WHY THIS EXISTS
 *
 * The structural checks were written against a PowerPoint: they read a slide's
 * text, its font sizes, its shape positions and its speaker notes. A PDF has none
 * of the latter three, so a student who supplied only a PDF got NO structural
 * checks at all — the app told them to add a .pptx.
 *
 * That is a real gap, because a PDF export is a perfectly normal way to hand in
 * work. This adapter supplies what CAN be known from a PDF — the text of each page,
 * the word counts, and how many images each page carries — so the text-based rules
 * can run. It deliberately does NOT invent the things a PDF cannot tell us: there
 * are no font sizes, no shape geometry and no notes, so the rules that need those
 * are skipped by lib/checks.js rather than being fed made-up numbers.
 *
 * The output mirrors the shape lib/pptx.js produces, so nothing downstream needs
 * to know which kind of file it came from.
 */

const fs = require('fs');
const crypto = require('crypto');
const { extractPdfText } = require('./pdftext');

/*
 * A slide's "title" in PowerPoint is a placeholder, which a PDF does not have.
 * The best available equivalent is the first line of the page when it is short,
 * which is what a heading looks like. This is a heuristic and is documented as
 * such — it only feeds the missing-title and title-too-long checks.
 */
const MAX_TITLE_CHARS = 90;
const MAX_TITLE_WORDS = 14;

function looksLikeTitle(line, isFirstLine) {
  if (!isFirstLine) return false;
  const t = String(line || '').trim();
  if (!t) return false;
  if (t.length > MAX_TITLE_CHARS) return false;
  if (t.split(/\s+/).length > MAX_TITLE_WORDS) return false;
  // A line ending in a full stop is prose, not a heading.
  if (/[.!?]$/.test(t)) return false;
  return true;
}

/**
 * Split a page's text into the word-count fields the checks use.
 *
 * `notesWordCount` is always 0 and that is correct, not a bug: a PDF export has no
 * speaker notes. The `missing-notes` rule is skipped for PDFs precisely so this
 * zero is never mistaken for "the student wrote no notes".
 */
function describePage(page, imageCount) {
  const text = String(page.text || '');
  const words = text.split(/\s+/).filter(Boolean);
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  const title = lines.length && looksLikeTitle(lines[0], true) ? lines[0] : null;

  /*
   * A single text block stands in for the page's text shapes. The checks that walk
   * `textShapes` are all geometry- or font-based and are skipped for PDFs, so this
   * exists only so nothing downstream throws on a missing array.
   */
  const textShape = {
    kind: 'body',
    name: 'Page text',
    isTitle: !!title,
    geometry: null,
    paragraphs: lines.map((l) => ({ text: l, sizePt: null, runs: [] })),
  };

  /*
   * A content fingerprint, matching what lib/pptx.js produces for a slide.
   *
   * The scanner uses this to notice which pages actually changed, so that repeat
   * checks can skip unchanged work. Without it every page would look "new" on every
   * scan and the run would never be incremental.
   */
  const fingerprint = crypto
    .createHash('sha1')
    .update(`${words.length}|${imageCount || 0}|${text}`)
    .digest('hex')
    .slice(0, 12);

  return {
    index: page.index,
    fingerprint,
    archivePath: null,
    title,
    hasTitle: !!title,
    hasTitlePlaceholder: !!title,
    text,
    textShapes: [textShape],
    wordCount: words.length,
    // The longest paragraph, which wall-of-text uses; harmless to supply.
    longestParagraphWords: lines.reduce((m, l) => Math.max(m, l.split(/\s+/).filter(Boolean).length), 0),
    images: [],
    imageCount: imageCount || 0,
    notesWordCount: 0,
    hasNotes: false,
    // Marks this page as PDF-derived, for anything that wants to know.
    fromPdf: true,
  };
}

/**
 * Build a deck object from a PDF buffer.
 *
 * @param {Buffer} buffer
 * @param {object} [opts] { fileName }
 * @returns {{ deck: object, warnings: string[] }}
 */
function deckFromPdf(buffer, opts = {}) {
  const parsed = extractPdfText(buffer);
  const warnings = (parsed && parsed.warnings) || [];
  const pages = (parsed && parsed.pages) || [];

  const slides = pages.map((p) => {
    // Accept either spelling from the extractor, so a small mismatch is not fatal.
    const imgs = p.hasImages != null ? p.hasImages : (p.imageCount || 0);
    return describePage(p, typeof imgs === 'number' ? imgs : (imgs ? 1 : 0));
  });

  const withText = slides.filter((s) => s.wordCount > 0).length;
  if (pages.length && !withText) {
    warnings.push('No text could be read from this PDF. If it is a scan or a set of '
      + 'photographs, the structural checks cannot read it — the AI visual review still can.');
  }

  /*
   * Page dimensions, taken from the PDF when the extractor reports them.
   *
   * These must be numbers rather than null: the report labels the size ("16:9
   * widescreen" etc.) and divides them, so a null crashes the whole scan at the end
   * — after all the work, which is the worst place to fail. 16:9 is the fallback
   * because that is what a PowerPoint/Slides export normally is, and a wrong label
   * is harmless whereas a crash is not.
   */
  const first = (parsed && parsed.pages && parsed.pages[0]) || {};
  const widthPt = Number(first.widthPt) > 0 ? Number(first.widthPt) : 960;
  const heightPt = Number(first.heightPt) > 0 ? Number(first.heightPt) : 540;
  const aspect = widthPt / heightPt;

  const deck = {
    source: 'pdf',
    fileName: opts.fileName || null,
    slideCount: slides.length,
    slides,
    slideWidthPt: widthPt,
    slideHeightPt: heightPt,
    slideWidthIn: widthPt / 72,
    slideHeightIn: heightPt / 72,
    aspect,
    notesSlideCount: 0,
    warnings,
  };
  return { deck, warnings };
}

function deckFromPdfFile(filePath, opts = {}) {
  // Read once and go through the same path as the buffer form, so the two cannot
  // diverge.
  return deckFromPdf(fs.readFileSync(filePath), opts);
}

module.exports = { deckFromPdf, deckFromPdfFile, looksLikeTitle, MAX_TITLE_CHARS };
