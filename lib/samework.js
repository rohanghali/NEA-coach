/*
 * NEA Coach — are the PowerPoint and the PDF the same portfolio?
 *
 * WHY THIS EXISTS
 *
 * Only ONE file is ever the structural deck — findDeck returns the .pptx if one is set, and the
 * PDF otherwise — but the PDF is ALSO the thing the AI looks at. So when both are set, the report
 * combines a layout analysis of one file with a visual mark of another. If they are not the same
 * piece of work, the report is about two portfolios at once and every number in it is wrong.
 *
 * The only check was the filename: "Rohan Ghali.pptx" and "Rohan Ghali.pdf" look like a pair, and
 * anything else raised a warning. That is a weak signal in both directions. A student who exports
 * to PDF gets whatever name the export tool chose, so a genuine pair can be flagged; and two
 * different portfolios called "Portfolio.pptx" and "Portfolio.pdf" sail straight through.
 *
 * WHAT THIS DOES INSTEAD
 *
 * Compares the WORDS. Both files are already read during a scan, and both parsers expose the text
 * of each page, so the evidence is in hand and costs one extra pass over the PDF.
 *
 * The measure is JACCARD — the words in BOTH, over the words in EITHER.
 *
 * Containment was the first choice, on the reasoning that a PDF export is often a subset of the
 * .pptx and containment would not punish it for being one. **Measuring killed it.** On this
 * project's own files, two DIFFERENT students' portfolios share 57%, 57% and 54% of their
 * distinctive words — every D&T portfolio is written in the same vocabulary of client, design,
 * material, specification and manufacture — against 99.7% for a portfolio and its own export.
 * Containment separated those two cases by less than the distance between two strangers, so a
 * threshold placed anywhere sensible would have called unrelated work the same.
 *
 * Jaccard on the same pairs: 0.98 for the genuine pair, 0.18 to 0.29 for the strangers. That is a
 * gap a threshold can sit inside.
 *
 * THRESHOLDS ARE MEASURED, NOT GUESSED — see the note above JACCARD_SAME. Nothing here decides
 * that work is different on a hunch: "unsure" is a real answer, and it is used generously.
 */

'use strict';

/*
 * Words too common to be evidence of anything. Without this, two unrelated documents share most
 * of their words simply because both are written in English, and the measure stops discriminating.
 */
const STOPWORDS = new Set([
  'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'any', 'can', 'had', 'her', 'was', 'one',
  'our', 'out', 'day', 'get', 'has', 'him', 'his', 'how', 'its', 'may', 'new', 'now', 'old', 'see',
  'two', 'who', 'did', 'yes', 'this', 'that', 'with', 'from', 'have', 'they', 'been', 'were', 'will',
  'your', 'when', 'what', 'into', 'than', 'then', 'them', 'some', 'such', 'only', 'over', 'also',
  'more', 'most', 'much', 'very', 'each', 'other', 'which', 'their', 'would', 'there', 'these',
  'about', 'been', 'could', 'should', 'because', 'between', 'through', 'during', 'before', 'after',
  'above', 'below', 'again', 'further', 'once', 'here', 'where', 'why', 'while', 'both', 'few',
  'being', 'does', 'doing', 'done', 'using', 'used', 'use', 'make', 'made', 'makes', 'can', 'may',
]);

/** The distinctive words of a page, or of a whole document flattened into one string. */
function tokensOf(text) {
  return String(text || '')
    .toLowerCase()
    // Anything that is not a letter or a digit becomes a break, so "CAD-modelled" and "CAD
    // modelled" tokenise together and punctuation cannot hide a match.
    .replace(/[^a-z0-9]+/g, ' ')
    .split(' ')
    // Three characters and up: this drops page numbers, "a", "of" and list bullets, all of which
    // are noise, while keeping "cad", "abs", "3d" and "5mm" as real evidence.
    .filter((w) => w.length >= 3 && !STOPWORDS.has(w));
}

/*
 * The two thresholds, and the measurement they came from.
 *
 *   a portfolio .pptx against its own PDF export   0.980   same work
 *   a PDF against itself                           1.000   same work
 *   Rohan's PDF against Arjun's PDF                0.286   different students
 *   Rohan's .pptx against Arjun's PDF              0.290   different students
 *   the bundled sample .pptx against Rohan's PDF   0.176   different work
 *
 * The gap runs from 0.29 to 0.98, so the bands sit either side of it. Everything in between
 * answers "unsure", which is a real answer and is used generously: a wrong "same" would let a
 * report quietly combine two pieces of work and present the total as one student's mark, which is
 * far worse than asking them to check.
 */
const JACCARD_SAME = 0.6;
const JACCARD_DIFFERENT = 0.35;
/*
 * Page counts are a second, independent signal, used for two things: to confirm a match, and to
 * create doubt. A genuine export differs by a page or two at most; the allowance is generous
 * because a .pptx can carry a slide the PDF export dropped.
 */
const PAGE_TOLERANCE = 0.25;
/* Below this there is not enough distinctive text to say anything at all. */
const MIN_TOKENS = 40;

/**
 * Compare two documents given their page texts.
 *
 * Shape-agnostic on purpose: it takes arrays of strings, so it depends on neither parser and can
 * be tested without either.
 *
 * @param {string[]} pptxPages text of each PowerPoint slide
 * @param {string[]} pdfPages text of each PDF page
 * @returns {object} verdict, the numbers behind it, and a sentence for the page
 */
function compareSameWork(pptxPages, pdfPages) {
  const pptxList = Array.isArray(pptxPages) ? pptxPages : [];
  const pdfList = Array.isArray(pdfPages) ? pdfPages : [];

  const pptxAll = tokensOf(pptxList.join(' '));
  const pdfAll = tokensOf(pdfList.join(' '));
  const pptxSet = new Set(pptxAll);
  const pdfSet = new Set(pdfAll);

  let shared = 0;
  for (const w of pptxSet) {
    if (pdfSet.has(w)) shared += 1;
  }

  const smaller = Math.min(pptxSet.size, pdfSet.size);
  const union = pptxSet.size + pdfSet.size - shared;
  const out = {
    verdict: 'unknown',
    reason: null,
    note: null,
    pptxPages: pptxList.length,
    pdfPages: pdfList.length,
    pptxWords: pptxSet.size,
    pdfWords: pdfSet.size,
    shared,
    // What share of the SMALLER document's distinctive words also appear in the larger.
    containment: smaller ? Math.round((shared / smaller) * 1000) / 1000 : 0,
    jaccard: union ? Math.round((shared / union) * 1000) / 1000 : 0,
  };

  if (smaller < MIN_TOKENS) {
    out.verdict = 'unknown';
    out.reason = 'not-enough-text';
    out.note = 'Could not compare these two files: one of them has almost no text to read, so there is nothing to match. That is normal for a portfolio that is mostly drawings.';
    return out;
  }

  const pages = Math.max(out.pptxPages, out.pdfPages);
  out.pageRatio = pages ? Math.abs(out.pptxPages - out.pdfPages) / pages : 0;
  const pagesAgree = out.pageRatio <= PAGE_TOLERANCE;
  const percent = Math.round(out.jaccard * 100);

  if (out.jaccard >= JACCARD_SAME && pagesAgree) {
    out.verdict = 'same';
    out.note = 'These two look like the same portfolio: ' + percent + '% of their distinctive words are common to both, and the page counts line up.';
    return out;
  }

  if (out.jaccard <= JACCARD_DIFFERENT) {
    out.verdict = 'different';
    out.note = 'These two do NOT look like the same portfolio. Only ' + percent + '% of their distinctive words are common to both, which is about what two unrelated pieces of work share. The report below would then mix a layout analysis of one with a visual mark of the other, and every number in it would be about neither one properly.';
    return out;
  }

  /*
   * Anything in between. Said plainly rather than rounded towards a verdict: this app's whole
   * problem has been numbers that look more certain than they are.
   */
  out.verdict = 'unsure';
  out.reason = pagesAgree ? 'partial-overlap' : 'page-counts-differ';
  out.note = pagesAgree
    ? percent + '% of their distinctive words are common to both. That is more than unrelated work but not enough to be sure, and the report combines the two — so check they really are the same portfolio.'
    : percent + '% of their distinctive words are common to both, and the page counts differ (' + out.pptxPages + ' against ' + out.pdfPages + '). Check they are the same portfolio, because the report combines them.';
  return out;
}

module.exports = {
  compareSameWork,
  tokensOf,
  STOPWORDS,
  JACCARD_SAME,
  JACCARD_DIFFERENT,
  PAGE_TOLERANCE,
  MIN_TOKENS,
};
