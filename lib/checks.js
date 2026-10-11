'use strict';
/*
 * checks.js — the free half of NEA Coach.
 *
 * These are deterministic, local rules: no AI, no network, no cost, and they
 * never invent a problem that is not really in the file. They catch the things
 * that genuinely lose marks on a portfolio, and they run in milliseconds.
 *
 * The AI pass (ai.js) handles judgement calls: "does this page actually
 * evidence the criterion?" This module handles the mechanical facts.
 */

/* ------------------------------------------------------------------ *
 * Thresholds, tuned for A3 portfolio pages.
 * A3 landscape is 16.54in x 11.69in — roughly 2x the area of A4 — so a page
 * can legitimately hold more than an A4 slide can.
 * ------------------------------------------------------------------ */
const THRESHOLDS = {
  // Words on one A3 page. A dense research page can be 400+; beyond ~520 the
  // page stops being readable in a moderation pile.
  wordsPerPageBusy: 400,
  wordsPerPageMax: 520,
  // One unbroken paragraph this long is very hard to read.
  wordsPerParagraphMax: 120,
  /*
   * The smallest body text a portfolio is allowed to use.
   *
   * This was 12pt. Lowered to 10 at the student's request: 11pt captions and
   * annotations were being reported as faults, and a rule that argues with work a
   * teacher has already accepted is a rule that gets ignored.
   *
   * One consequence is worth knowing. The bundled sample's planted "text too small"
   * flaw is an 11pt page, so the sample is now deliberately CLEAN for this rule and
   * no longer demonstrates it. Section 3 of the tests asserts that, and section 3b
   * carries the rule's real coverage from synthetic decks either side of the limit.
   */
  minFontPt: 10,
  // Smallest body text on the page.
  minBodyFontPt: 14,
  /*
   * IMAGE QUALITY THRESHOLDS - calibrated against real marked work.
   *
   * These were originally set for print perfectionism (flag anything under
   * 150 DPI) and that produced a wall of false alarms. Calibrated against a
   * 60-page portfolio that was awarded 99/100 at GCSE, the distribution was:
   *   200 images above 300 DPI, 100 at 150-300, 39 below 150, worst 70.
   * So sub-150 DPI images are NORMAL and acceptable in a student portfolio -
   * they are usually web-sourced reference photos or photographed sketchbook
   * pages. Flagging them as defects buried the real problems.
   *
   * Only genuinely unreadable images are worth raising now, and the opposite
   * fault matters too: a high-resolution image shrunk to a postage stamp is
   * unreadable for a completely different reason.
   */
  minImageDpi: 72, // below this it is visibly blocky, not merely soft
  smallImageIn: 1.2, // placed narrower than this, an image is hard to read
  smallImageDpi: 300, // ...unless it is also low-res, which is a different complaint
  // A page with fewer than this many words is probably carrying no content.
  nearEmptyWords: 12,
  // Title length sanity.
  maxTitleWords: 12,
  // PowerPoint canvas is 1pt larger than the slide for each side.
  canvasSlackPt: 2,
};

/* ------------------------------------------------------------------ *
 * Finding construction
 * ------------------------------------------------------------------ */

const SEVERITY_ORDER = { critical: 0, high: 1, medium: 2, low: 3, info: 4 };

let findingSeq = 0;
function makeFinding(f, ctx) {
  findingSeq += 1;
  const slide = ctx.slideNumber != null ? ctx.slideNumber : null;
  return {
    id: `${f.rule}@${slide == null ? 'deck' : slide}#${findingSeq}`,
    rule: f.rule,
    title: f.title,
    severity: f.severity,
    slide,
    slides: f.slides || (slide != null ? [slide] : []),
    location: f.location || null,
    evidence: f.evidence || null,
    why: f.why,
    fix: f.fix,
    effort: f.effort || 'minutes',
    source: 'local',
    tier: tierOf(f.rule),
    criteria: f.criteria || [],
    /*
     * Priority: how much the item deserves attention, then severity, then how
     * many pages it affects. The tier term is deliberately large enough to
     * outrank severity: a 'high' font complaint must never push a 'high'
     * reasoning gap down the list, because the student can already see the font.
     */
    priority: (TIERS[tierOf(f.rule)] || 2) * 1000
      + ({ critical: 100, high: 70, medium: 40, low: 15, info: 5 }[f.severity] || 0)
      + Math.min(20, (f.slides ? f.slides.length : 1) * 2),
  };
}

/* ------------------------------------------------------------------ *
 * Text overflow estimate
 * ------------------------------------------------------------------ */

/**
 * Estimate whether text overflows its text box.
 * We use an average glyph advance of 0.52em for typical sans-serif body text,
 * then allow for line spacing. This is an estimate by design: it is a useful
 * early warning, not a measurement. The AI/visual pass confirms real overflow.
 */
function estimateOverflow(shape, slideWidthPt, slideHeightPt) {
  if (!shape.geometry || !shape.paragraphs || !shape.paragraphs.length) return null;
  const { wPt, hPt } = shape.geometry;
  if (!wPt || !hPt || wPt <= 0 || hPt <= 0) return null;

  let neededPt = 0;
  let widest = 0;
  for (const p of shape.paragraphs) {
    const size = p.sizePt || 18;
    const text = p.text || '';
    if (!text.trim()) {
      neededPt += size * 1.2;
      continue;
    }
    const charWidth = size * 0.52;
    const perLine = Math.max(1, Math.floor(wPt / charWidth));
    const lines = Math.max(1, Math.ceil(text.length / perLine));
    neededPt += lines * size * 1.25;
    widest = Math.max(widest, Math.min(1, text.length / Math.max(1, perLine)));
  }

  const ratio = neededPt / hPt;
  return {
    neededPt: Math.round(neededPt),
    boxPt: Math.round(hPt),
    ratio: Number(ratio.toFixed(2)),
    overflows: ratio > 1.02,
    longestLineFill: Number(widest.toFixed(2)),
  };
}

/* ------------------------------------------------------------------ *
 * Rules
 * ------------------------------------------------------------------ */

/*
 * How much each rule deserves the student's attention.
 *
 * The student can already SEE whether a picture sits off the edge of the page
 * or whether a font is small. Listing 60 of those buries the handful of items
 * that actually cost marks — weak justification, description where analysis is
 * needed, informal wording, a claim with no evidence. So every rule carries a
 * tier and the app leads with 'substance'.
 *
 *   substance    — needs the student to think, redraft or add reasoning.
 *   presentation — worth tidying, but visible on the page.
 *   obvious      — mechanical; the student can spot these unaided.
 */
const TIERS = { substance: 3, presentation: 2, obvious: 1 };

const RULE_TIER = {
  'missing-title': 'presentation',
  'missing-notes': 'presentation',
  'thin-visual-evidence': 'substance',
  'duplicate-pages': 'substance',
  'near-empty-page': 'substance',
  'wall-of-text': 'presentation',
  'terminology-inconsistency': 'substance',
  'common-misspellings': 'substance',
  'font-inconsistency': 'presentation',
  'title-too-long': 'presentation',
  'deck-size': 'obvious',
  'page-too-busy': 'obvious',
  'text-too-small': 'obvious',
  'text-shrunk-to-fit': 'obvious',
  'off-canvas': 'obvious',
  'text-overflow': 'obvious',
  'uncaptioned-image': 'obvious',
  'low-res-image': 'obvious',
  'image-too-small': 'obvious',
};

/*
 * Rules switched off entirely.
 *
 * The student asked for the resolution checks to go: "can you make it not check
 * for image resolution because i can see that". They were also the two noisiest
 * rules in the set — 8 of 60 pages flagged for a picture sitting slightly under
 * 300 DPI. Whether a picture is good enough is a human call, so the
 * measurements are still parsed and reported per image, but they are no longer
 * raised as findings.
 */
const DISABLED_RULES = new Set(['low-res-image', 'image-too-small']);

const RULES = [];

/*
 * A rule may require something only a PowerPoint has.
 *
 * A PDF has no font sizes, no shape geometry and no speaker notes, so rules built
 * on those cannot run on one. Rather than let them quietly report "nothing found"
 * — which reads as a pass — each is tagged with what it needs, and a PDF run lists
 * the ones it could not do.
 *
 *   'text'     works from page text alone (word counts, spelling, duplication)
 *   'images'   needs the page's images to be identifiable
 *   'geometry' needs shape positions and sizes
 *   'fonts'    needs font sizes
 *   'notes'    needs speaker notes, which a PDF export does not contain
 */
function rule(name, fn, requires) {
  RULES.push({ name, fn, requires: requires || 'text' });
}

function tierOf(name) {
  return RULE_TIER[name] || 'presentation';
}

/* ---------- per-slide rules ---------- */

rule('missing-title', (deck, out) => {
  for (const s of deck.slides) {
    if (s.hasTitle) continue;
    out.push(makeFinding({
      rule: 'missing-title',
      title: 'Page has no clear title',
      severity: 'high',
      why: 'A moderator needs to know in one glance what each page is for. Untitled pages force them to read everything to find out, and they are the single most common reason a portfolio feels disorganised.',
      fix: 'Add a short, specific heading at the top of the page, e.g. "Research: existing products" rather than just "Research".',
      evidence: s.wordCount ? `Page has ${s.wordCount} words but the largest text block is not a heading.` : 'Page has no substantial text at all.',
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
});

rule('page-too-busy', (deck, out) => {
  for (const s of deck.slides) {
    if (s.wordCount < THRESHOLDS.wordsPerPageBusy) continue;
    const severity = s.wordCount >= THRESHOLDS.wordsPerPageMax ? 'high' : 'medium';
    out.push(makeFinding({
      rule: 'page-too-busy',
      title: `Page is overloaded with text (${s.wordCount} words on one A3 page)`,
      severity,
      why: 'Marks are awarded for how clearly you communicate, not for volume. When a page holds this much prose the reader skips it, and any evidence buried inside it stops counting towards your mark.',
      fix: 'Split this into two pages, or move supporting detail into annotated photographs and a short table. Aim for one clear idea per page.',
      evidence: `${s.wordCount} words across ${s.textShapes.length} text block(s). Longest single paragraph: ${s.longestParagraphWords} words.`,
      effort: 'hours',
    }, { slideNumber: s.index }));
  }
}, 'geometry');

rule('wall-of-text', (deck, out) => {
  for (const s of deck.slides) {
    for (const shape of s.textShapes) {
      if (shape.isTitle) continue;
      for (const p of shape.paragraphs) {
        const words = (p.text || '').split(/\s+/).filter(Boolean).length;
        if (words < THRESHOLDS.wordsPerParagraphMax) continue;
        out.push(makeFinding({
          rule: 'wall-of-text',
          title: `Single unbroken paragraph of ${words} words`,
          severity: 'medium',
          why: 'A block of text this long is read as a wall and skipped. Examiners respond far better to short paragraphs, bullet points, or a labelled diagram with a caption.',
          fix: 'Break it into 3-5 bullet points, or turn the same information into a labelled sketch or table. Keep the detail, change the presentation.',
          evidence: p.text.slice(0, 140) + (p.text.length > 140 ? '…' : ''),
          effort: 'minutes',
        }, { slideNumber: s.index }));
      }
    }
  }
}, 'geometry');

rule('text-too-small', (deck, out) => {
  for (const s of deck.slides) {
    const small = [];
    for (const shape of s.textShapes) {
      if (shape.isTitle) continue;
      for (const p of shape.paragraphs) {
        const size = p.sizePt != null ? p.sizePt : (shape.sizes && Object.keys(shape.sizes).length
          ? Number(Object.keys(shape.sizes)[0])
          : null);
        if (size == null) continue;
        if (size < THRESHOLDS.minFontPt) {
          small.push({ size, words: (p.text || '').split(/\s+/).filter(Boolean).length });
        }
      }
    }
    if (!small.length) continue;
    const minSize = Math.min(...small.map((x) => x.size));
    out.push(makeFinding({
      rule: 'text-too-small',
      title: `Body text drops to ${minSize}pt`,
      /*
       * Tied to the threshold rather than to a hard-coded 10. It used to read
       * "minSize <= 10", which silently became "always high" the instant the minimum
       * was lowered to 10 — the rule would have lost its severity scale without
       * anything failing. Two points below the limit is the same margin it always
       * used.
       */
      severity: minSize <= THRESHOLDS.minFontPt - 2 ? 'high' : 'medium',
      why: 'Your portfolio will be read on screen and possibly printed. Text this small is a strain on an A3 sheet, and a marker who has to squint reads less of your work.',
      fix: `Raise body text to at least ${THRESHOLDS.minBodyFontPt}pt. If that makes the page too full, the page is carrying too much — split it.`,
      evidence: `${small.length} paragraph(s) below ${THRESHOLDS.minFontPt}pt; smallest is ${minSize}pt.`,
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
}, 'fonts');

rule('low-res-image', (deck, out) => {
  for (const s of deck.slides) {
    const offenders = s.images.filter(
      (im) => !im.isVector && im.effectiveDpi != null && im.effectiveDpi < THRESHOLDS.minImageDpi
    );
    if (!offenders.length) continue;
    const worst = offenders.reduce((a, b) => (a.effectiveDpi < b.effectiveDpi ? a : b));

    /*
     * Only genuinely blocky images are worth a student's time. Anything softer
     * than that is normal for a photographed sketchbook page, and saying so at
     * low severity keeps the report honest without burying the real problems.
     */
    const blocky = worst.effectiveDpi < THRESHOLDS.minImageDpi;
    out.push(makeFinding({
      rule: 'low-res-image',
      title: worst.effectiveDpi < 50
        ? `Image is too low-resolution to read (${worst.effectiveDpi} DPI)`
        : `One image is noticeably soft (${worst.effectiveDpi} DPI effective)`,
      severity: worst.effectiveDpi < 50 ? 'medium' : 'low',
      why: 'This is a polish point, not a mark-loser. A photographed page or a web image shown large will look slightly soft on screen and in print. It only becomes a real problem if a marker cannot make out the detail you are pointing at.',
      fix: worst.isVector
        ? 'Re-export this as a vector if you still have the original file.'
        : `If you still have the original, re-export it larger. Otherwise leave it: this image is ${worst.pixelWidth}x${worst.pixelHeight}px shown at ${worst.placedWidthIn}in wide, which is fine to read.`,
      evidence: `${offenders.length} image(s) below ${THRESHOLDS.minImageDpi} DPI. Lowest: ${worst.pixelWidth}x${worst.pixelHeight}px shown at ${worst.placedWidthIn}in wide (${worst.effectiveDpi} DPI).`,
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
}, 'images');

/**
 * The opposite fault: an image with plenty of pixels squeezed so small that its
 * detail cannot be read at all. On the 99/100 portfolio this affected several
 * images placed under an inch wide, which the old low-resolution rule ignored
 * precisely because their DPI looked excellent.
 */
rule('image-too-small', (deck, out) => {
  for (const s of deck.slides) {
    const tiny = s.images.filter(
      (im) =>
        im.placedWidthIn != null &&
        im.placedWidthIn > 0 &&
        im.placedWidthIn < THRESHOLDS.smallImageIn &&
        (im.effectiveDpi == null || im.effectiveDpi >= THRESHOLDS.smallImageDpi)
    );
    if (!tiny.length) continue;
    const worst = tiny.reduce((a, b) => (a.placedWidthIn < b.placedWidthIn ? a : b));
    out.push(makeFinding({
      rule: 'image-too-small',
      title: `${tiny.length} image${tiny.length > 1 ? 's are' : ' is'} too small to read (${worst.placedWidthIn}in wide)`,
      severity: 'low',
      why: 'A marker has to be able to see what your photograph or drawing shows. At this size the image reads as texture rather than evidence, so it earns nothing even though it is technically present.',
      fix: 'Make the image bigger, or crop it to the part that matters so the detail fills the space. One large, readable photo beats four thumbnails.',
      evidence: `Smallest: ${worst.pixelWidth}x${worst.pixelHeight}px displayed only ${worst.placedWidthIn}in wide (${worst.effectiveDpi} DPI — plenty of pixels, just no size).`,
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
}, 'geometry');

rule('uncaptioned-image', (deck, out, ctx) => {
  /*
   * Some boards do not require annotation at all.
   *
   * WJEC/Eduqas says plainly that "annotation throughout the work is not
   * necessary". Telling an Eduqas student to caption 25 pages is therefore advice
   * against their own board's rules, and on a real portfolio it was 25 of 103
   * findings — noise that buried the advice that mattered. The rule still runs for
   * boards that do want annotation.
   */
  if (ctx && ctx.annotationRequired === false) return;
  for (const s of deck.slides) {    if (!s.images.length) continue;

    /*
     * A page with almost no text at all is already reported as a near-empty
     * page, which is the more serious problem. Reporting its images as
     * uncaptioned too would double-count the same fault and pad the list.
     */
    if (s.wordCount < THRESHOLDS.nearEmptyWords) continue;

    /*
     * Decide PER IMAGE whether it carries its own annotation.
     *
     * An earlier version skipped any page with more than 60 words, on the theory
     * that the prose probably explained the picture. That made the rule inert on
     * every real portfolio page, because real pages always carry prose. It only
     * looked like it worked on the test deck, whose pages were unusually short.
     *
     * The second version was still too narrow: it only looked BELOW each image.
     * Rohan's portfolio captions its images BESIDE them, so images paired with a
     * clear label sitting next to them were all reported as unannotated, which
     * produced 41 noisy findings at once. A caption counts if it falls anywhere
     * in a band covering the image and a margin above and below.
     */
    const unannotated = [];
    for (const im of s.images) {
      if (im.yIn == null) continue;
      const h = im.placedHeightIn || 1;
      const bandTop = im.yIn - Math.max(0.6, h * 0.35);
      const bandBottom = im.yIn + h + Math.max(0.8, h * 0.35);

      const hasCaption = s.textShapes.some((sh) => {
        if (sh.isTitle || !sh.geometry || !sh.paragraphs) return false;
        const w = sh.wordCount;
        if (!w || w > 90) return false; // a whole body column is not a caption
        const g = sh.geometry;
        if (g.yIn == null) return false;
        if (g.yIn < bandTop || g.yIn > bandBottom) return false;
        /*
         * If the text block is wide enough it is page prose that happens to
         * start nearby, not a caption attached to this image. Narrow blocks and
         * ones on the same horizontal side as the image are treated as captions.
         */
        const wide = g.wIn != null && g.wIn > 7.5;
        if (wide) return false;
        return true;
      });

      if (!hasCaption) unannotated.push(im);
    }

    if (!unannotated.length) continue;

    const names = unannotated.map((im) => `${im.shapeName || 'image'} at ${im.placedWidthIn}x${im.placedHeightIn}in`);
    out.push(makeFinding({
      rule: 'uncaptioned-image',
      title: unannotated.length === s.images.length
        ? `Image with no caption or annotation (${unannotated.length} image${unannotated.length > 1 ? 's' : ''})`
        : `${unannotated.length} of ${s.images.length} images have no caption or annotation`,
      severity: 'medium',
      why: 'An unannotated image earns almost nothing. Markers are looking for your thinking, and a caption is where you say what the image proves, what worked, or what you would change.',
      fix: 'Add a one-line caption under each image saying what it shows and what it tells you. "Figure 4: prototype 2 — the 6mm ply split along the grain when clamped, so I changed to 12mm."',
      evidence: names.join(', '),
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
}, 'geometry');

rule('near-empty-page', (deck, out) => {
  for (const s of deck.slides) {
    if (s.wordCount >= THRESHOLDS.nearEmptyWords) continue;
    /*
     * How many images are on this page, whichever source it came from.
     *
     * This guard has always existed — "a big annotated picture page can be fine" —
     * but it read `s.images.length`, and the PDF path never populates image
     * OBJECTS, only a count. So on a PDF the guard was always 0 and every
     * image-heavy page was judged on words alone. On a real AQA portfolio that
     * scored 90/100 it flagged the cover as "carries only 4 words", which is a
     * high-severity finding about a page that is entirely correct.
     */
    const imageCount = typeof s.imageCount === 'number' ? s.imageCount : (s.images ? s.images.length : 0);
    if (imageCount > 0 && s.wordCount >= 4) continue;
    // A cover or divider legitimately carries almost nothing. Only a substantial
    // portfolio has one, so a two-page deck is still checked in full.
    if (imageCount > 0 && deck.slideCount >= 4) continue;
    out.push(makeFinding({
      rule: 'near-empty-page',
      title: s.wordCount === 0 ? 'Page is completely blank of text' : `Page carries only ${s.wordCount} words`,
      severity: 'high',
      why: 'Every page in a portfolio should either present evidence or explain a decision. A near-empty page reads as unfinished work and uses up a page of a 200-page submission for nothing.',
      fix: 'Either develop this page properly, or delete it. Do not pad it with filler text — a shorter, denser portfolio scores better than a long empty one.',
      evidence: s.wordCount === 0 ? 'No text found on this page.' : `Only: "${(s.text || '').slice(0, 90)}"`,
      effort: 'hours',
    }, { slideNumber: s.index }));
  }
});

rule('missing-notes', (deck, out) => {
  const missing = deck.slides.filter((s) => s.notesWordCount === 0 && s.wordCount > 0);
  // Only flag if the deck uses notes at all — otherwise it is a style choice.
  const pagesWithNotes = deck.slides.filter((s) => s.notesWordCount > 0).length;
  if (pagesWithNotes < Math.max(3, deck.slideCount * 0.15)) return; // deck does not use notes; skip
  
  for (const s of missing) {
    out.push(makeFinding({
      rule: 'missing-notes',
      title: 'No speaker notes on this page',
      severity: 'low',
      why: 'You already use notes elsewhere in this deck. Notes are your record of the reasoning behind each page, and they are the easiest place to show evaluation and justification without cluttering the page itself.',
      fix: 'Add two or three lines in the notes box: what decision you made on this page, why, and what you would do differently.',
      evidence: `Notes are used on ${pagesWithNotes} other page(s) but not here.`,
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
}, 'notes');

rule('text-overflow', (deck, out) => {
  for (const s of deck.slides) {
    for (const shape of s.textShapes) {
      /*
       * A table is not a fixed box: PowerPoint grows rows to fit their content,
       * so the recorded frame height is not a clipping boundary. Estimating
       * overflow from it produced instant false positives on every table page
       * once table text started being extracted.
       */
      if (shape.kind === 'table') continue;

      const est = estimateOverflow(shape, deck.slideWidthPt, deck.slideHeightPt);
      if (!est || !est.overflows) continue;

      /*
       * Honour autofit. This is the difference between a real defect and a
       * false alarm, and the distinction was verified against an actual render:
       *
       *  - 'shape-grows'  (<a:spAutoFit/>): the box expands to fit the text, so
       *    text beyond the recorded height is NORMAL. The only real problem is
       *    if the grown box would run off the bottom of the page.
       *  - 'text-shrinks' (<a:normAutofit/>): the text is scaled down to fit, so
       *    nothing is clipped, but it may have been shrunk to unreadable.
       *  - 'none': a fixed box, where overflow genuinely means clipped text.
       */
      if (shape.autofit === 'shape-grows') {
        if (!shape.geometry || shape.geometry.yPt == null) continue;
        const overflowPt = shape.geometry.yPt + est.neededPt - deck.slideHeightPt;
        if (overflowPt <= 24) continue; // absorbed by the page margin
        out.push(makeFinding({
          rule: 'text-overflow',
          title: `Box grows past the bottom of the page (about ${Math.round(overflowPt)}pt over)`,
          severity: overflowPt > 120 ? 'high' : 'medium',
          why: 'This text box resizes itself to fit its content, and with this much text it expands beyond the page edge. The last lines get cut off when the portfolio is printed or exported to PDF.',
          fix: 'Cut some text, move it onto its own page, or drag the box up. Keeping to one clear idea per page stops this happening.',
          evidence: `Content needs about ${est.neededPt}pt starting at ${Math.round(shape.geometry.yPt)}pt, but the page is only ${Math.round(deck.slideHeightPt)}pt tall.`,
          effort: 'minutes',
        }, { slideNumber: s.index }));
        continue;
      }

      if (shape.autofit === 'text-shrinks') {
        // fontScale is stored in thousandths of a percent, e.g. 62500 = 62.5%.
        const scale = shape.fontScale != null ? shape.fontScale / 100000 : null;
        const firstSized = shape.paragraphs.find((p) => p.sizePt);
        const baseSize = (firstSized && firstSized.sizePt) || 18;
        const effective = scale != null ? baseSize * scale : baseSize;
        if (scale == null || effective >= THRESHOLDS.minFontPt) continue;
        out.push(makeFinding({
          rule: 'text-shrunk-to-fit',
          title: `Text has been auto-shrunk to about ${effective.toFixed(1)}pt to fit the box`,
          severity: effective < 9 ? 'high' : 'medium',
          why: 'PowerPoint is shrinking your text to squeeze it into this box. It looks acceptable on screen, but on paper or projected it becomes too small to read, and unread evidence earns nothing.',
          fix: 'Reduce the amount of text on this page instead of letting PowerPoint shrink it. Split the page, or move the detail into the speaker notes.',
          evidence: `Box reports a font scale of ${(scale * 100).toFixed(0)}% applied to ${baseSize}pt text.`,
          effort: 'minutes',
        }, { slideNumber: s.index }));
        continue;
      }

      out.push(makeFinding({
        rule: 'text-overflow',
        title: `Text likely spills outside its box (needs about ${est.neededPt}pt, box is ${est.boxPt}pt)`,
        severity: est.ratio > 1.5 ? 'high' : 'medium',
        why: 'This box has a fixed size, so text that does not fit is clipped when printed, or overlaps whatever sits below it. Either way the marker cannot read it, so those marks are lost.',
        fix: 'Shrink the paragraph, enlarge the box, or move the overflow onto its own page. Then check the page at print size to confirm.',
        evidence: `Estimated content height ${est.neededPt}pt in a fixed box of ${est.boxPt}pt (${Math.round(est.ratio * 100)}%).`,
        effort: 'minutes',
      }, { slideNumber: s.index }));
    }
  }
}, 'geometry');

rule('off-canvas', (deck, out) => {
  const w = deck.slideWidthPt;
  const h = deck.slideHeightPt;
  for (const s of deck.slides) {
    const strays = [];
    for (const sh of s.shapes) {
      const g = sh.geometry;
      if (!g || g.xPt == null || g.yPt == null) continue;
      const slack = THRESHOLDS.canvasSlackPt;
      if (g.xPt < -slack || g.yPt < -slack || g.xPt + g.wPt > w + slack || g.yPt + g.hPt > h + slack) {
        strays.push(sh.name || sh.kind);
      }
    }
    if (!strays.length) continue;
    out.push(makeFinding({
      rule: 'off-canvas',
      title: `${strays.length} element(s) sit outside the printable page area`,
      severity: 'medium',
      why: 'Anything outside the page boundary disappears when the portfolio is printed or exported to PDF, which is how your work will be viewed.',
      fix: 'Drag these elements back inside the page, or reduce their size. Check the page in print preview as well.',
      evidence: `Outside the page: ${strays.slice(0, 5).join(', ')}${strays.length > 5 ? ` (+${strays.length - 5} more)` : ''}.`,
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
}, 'geometry');

rule('thin-visual-evidence', (deck, out) => {
  const textOnly = deck.slides.filter((s) => s.imageCount === 0 && s.wordCount > 40);
  if (textOnly.length < 3) return;
  out.push(makeFinding({
    rule: 'thin-visual-evidence',
    title: `${textOnly.length} text-heavy pages contain no images at all`,
    severity: 'medium',
    slides: textOnly.map((s) => s.index),
    why: 'This is a design and technology portfolio. Sketches, photographs of models and annotated screenshots are the primary evidence of your making and iteration, and they carry marks that prose alone cannot.',
    fix: 'For each of these pages add at least one piece of visual evidence: a photograph of the prototype, a sketch, a dimensioned drawing, or a screenshot of test results. Annotate it.',
    evidence: `Pages with over 40 words and no images: ${textOnly.map((s) => s.index).join(', ')}.`,
    effort: 'hours',
  }, {}));
});

/* ---------- deck-level rules ---------- */

rule('duplicate-pages', (deck, out) => {
  const byContent = new Map();
  for (const s of deck.slides) {
    const key = (s.text || '').replace(/\s+/g, ' ').trim().toLowerCase();
    if (key.length < 40) continue;
    if (!byContent.has(key)) byContent.set(key, []);
    byContent.get(key).push(s.index);
  }
  for (const [, indexes] of byContent) {
    if (indexes.length < 2) continue;
    out.push(makeFinding({
      rule: 'duplicate-pages',
      title: `Duplicate content on pages ${indexes.join(' and ')}`,
      severity: 'high',
      slides: indexes,
      why: 'Repeated pages look like an accident and waste space in a submission where every page should earn something. If you meant to show development, the second version must be visibly different.',
      fix: 'Delete the duplicate, or change it so it clearly shows progression — label it "version 1" and "version 2" and say what changed and why.',
      evidence: `Identical text content found on ${indexes.length} pages.`,
      effort: 'minutes',
    }, {}));
  }
}, 'text');

rule('font-inconsistency', (deck, out) => {
  /*
   * Which lines count as headings is a judgement, and on a PDF it is made from
   * font size alone — there are no placeholder roles to consult. Compared with a
   * document's own typical body size it is reliable, so anything only slightly
   * larger than the body text is treated as body text rather than as a heading
   * of its own size. Without this the rule fires on the small print of a single
   * page and calls it a typography problem.
   */
  const allSizes = [];
  for (const s of deck.slides) {
    for (const sh of s.textShapes) {
      for (const p of sh.paragraphs) {
        const n = p.sizePt != null ? Number(p.sizePt) : null;
        if (n && isFinite(n) && n > 0) allSizes.push(n);
      }
    }
  }
  let bodySize = 0;
  if (allSizes.length) {
    allSizes.sort((a, b) => a - b);
    bodySize = allSizes[Math.floor(allSizes.length / 2)];
  }
  const headingFloor = bodySize > 0 ? Math.max(12, bodySize * 1.4) : 12;

  /*
   * Group sizes that are the same size in design terms.
   *
   * A measured size is never exactly what the designer typed: PDF text is scaled
   * by the page transform, so one 32pt heading arrives as 32 and another as 32.1,
   * and a 14pt one as 14, 14.0 and 14.2. Counting those as separate heading sizes
   * produced "3 different sizes" for a portfolio using two, on a real AQA
   * portfolio that scored 90/100. Sizes within 2% (and at least a third of a
   * point, so tiny text is not merged) are treated as one.
   */
  const sameSize = (a, b) => Math.abs(a - b) <= Math.max(0.35, Math.max(a, b) * 0.02);

  const titleSizes = new Map();
  for (const s of deck.slides) {
    for (const sh of s.textShapes) {
      if (!sh.isTitle) continue;
      for (const p of sh.paragraphs) {
        const size = p.sizePt != null ? Number(p.sizePt) : null;
        if (size == null || size < headingFloor) continue;
        // Fold this size into an existing cluster, or start a new one.
        let matched = null;
        for (const key of titleSizes.keys()) {
          if (sameSize(key, size)) { matched = key; break; }
        }
        const bucket = matched != null ? matched : size;
        titleSizes.set(bucket, (titleSizes.get(bucket) || 0) + 1);
      }
    }
  }
  if (titleSizes.size <= 2) return;
  const sizes = Array.from(titleSizes.keys()).sort((a, b) => a - b);
  out.push(makeFinding({
    rule: 'font-inconsistency',
    title: `Headings use ${titleSizes.size} different sizes across the deck`,
    severity: 'low',
    slides: deck.slides.map((s) => s.index),
    why: 'Consistent typography is the cheapest way to make a long document feel designed rather than assembled. Wobbling heading sizes is the first thing a moderator notices.',
    fix: `Pick one heading size and apply it to every page. Sizes currently in use: ${sizes.join('pt, ')}pt.`,
    evidence: `Heading sizes found: ${sizes.map((sz) => `${sz}pt x${titleSizes.get(sz)}`).join(', ')}.`,
    effort: 'hours',
  }, {}));
}, 'fonts');

rule('terminology-inconsistency', (deck, out) => {
  /*
   * Pairs where ONE form is British and the other is American.
   *
   * Two pairs were removed from this list because they are not misspellings at all, and the rule
   * was telling students to change correct English:
   *
   *   metre / meter   — a metre is the unit; a METER is a measuring instrument. Both British.
   *   licence / license — LICENCE is the noun, LICENSE is the verb. Both British ("poetic licence",
   *                       "licensed to..."). This is the classic British pair, not an error.
   *
   * The patterns are also anchored at BOTH ends now. Without the trailing boundary, \bcolor matched
   * "colorimeter" and \bmeter matched "parameter" and "diameter", so a student writing about a
   * colorimeter was accused of mixing in American spelling they had never used. A rule that invents
   * a fault costs more trust than the fault it found.
   */
  const PAIRS = [
    ['aluminium', 'aluminum'],
    ['colour', 'color'],
    ['centre', 'center'],
    ['modelling', 'modeling'],
    ['fibre', 'fiber'],
  ];
  for (const [uk, us] of PAIRS) {
    const hasUk = deck.slides.some((s) => new RegExp(`\\b${uk}\\b`, 'i').test(s.text));
    const hasUs = deck.slides.some((s) => new RegExp(`\\b${us}\\b`, 'i').test(s.text));
    if (!(hasUk && hasUs)) continue;
    const slides = deck.slides.filter((s) => new RegExp(`\\b(${uk}|${us})\\b`, 'i').test(s.text)).map((s) => s.index);
    out.push(makeFinding({
      rule: 'terminology-inconsistency',
      title: `Spelling switches between "${uk}" and "${us}"`,
      severity: 'low',
      slides,
      why: 'Inconsistent spelling reads as carelessness and, in a technical portfolio, inconsistent terminology can make it unclear whether you mean two different things.',
      fix: `Standardise on "${uk}" throughout (UK English is expected by UK exam boards), and keep to one term per material or process.`,
      evidence: `Both forms appear across pages ${slides.join(', ')}.`,
      effort: 'minutes',
    }, {}));
  }
});

/*
 * Common misspellings, and the words most often confused with them.
 *
 * Two separate lists on purpose, because they are different kinds of claim:
 *
 *   TYPOS      — always wrong. Safe to state flatly.
 *   CONFUSED   — both words are real; only the writer knows which was meant. The
 *                finding says "check this", never "this is wrong", because
 *                telling a student to change "effect" to "affect" when they meant
 *                "effect" is exactly the kind of false correction that destroys
 *                trust in every other finding.
 *
 * Technical and D&T terms are deliberately ABSENT, even where they look like
 * typos ("millimetres" is correct British spelling; "aluminium" is not a typo for
 * "aluminum"). A spell list that flags the student's own subject vocabulary is
 * worse than no spell list.
 */
const TYPOS = [
  // Clear-cut only. Anything arguable is NOT here: flagging a correct word is
  // worse than missing a typo, because it teaches the student to distrust the
  // list. Removed for that reason: "dependant" (correct British English),
  // "publically" (an accepted variant), and "medieval"/"millimetres" (correct).
  ['recieved', 'received'], ['seperate', 'separate'], ['definately', 'definitely'],
  ['occured', 'occurred'], ['wich', 'which'], ['thier', 'their'],
  ['accomodate', 'accommodate'], ['saftey', 'safety'], ['choosen', 'chosen'],
  ['sucessful', 'successful'], ['strenghts', 'strengths'], ['compontent', 'component'],
  ['enviroment', 'environment'], ['goverment', 'government'], ['experiance', 'experience'],
  ['beleive', 'believe'], ['freind', 'friend'], ['wierd', 'weird'], ['untill', 'until'],
  ['begining', 'beginning'], ['commitee', 'committee'], ['neccessary', 'necessary'],
  ['recomend', 'recommend'], ['tommorow', 'tomorrow'],
  ['arguement', 'argument'], ['maintainance', 'maintenance'], ['managment', 'management'],
  ['occassion', 'occasion'], ['posession', 'possession'],
  ['reccomend', 'recommend'], ['refered', 'referred'], ['relevent', 'relevant'],
  ['resistence', 'resistance'], ['responsable', 'responsible'], ['seperately', 'separately'],
  ['similiar', 'similar'], ['suprise', 'surprise'], ['tendancy', 'tendency'],
  ['acheive', 'achieve'], ['adress', 'address'], ['agressive', 'aggressive'],
  ['apparant', 'apparent'], ['basicly', 'basically'], ['calender', 'calendar'],
  ['catagory', 'category'], ['completly', 'completely'], ['concensus', 'consensus'],
  ['desicion', 'decision'], ['diffrent', 'different'], ['dissapoint', 'disappoint'],
  ['embarass', 'embarrass'], ['equiptment', 'equipment'], ['existance', 'existence'],
  ['familar', 'familiar'], ['finaly', 'finally'], ['foriegn', 'foreign'],
  ['garentee', 'guarantee'], ['happend', 'happened'],
  ['immediatly', 'immediately'], ['independant', 'independent'], ['knowlege', 'knowledge'],
  ['liason', 'liaison'], ['libary', 'library'],
];


rule('common-misspellings', (deck, out) => {
  for (const [wrong, right] of TYPOS) {
    const hits = deck.slides.filter((s) => new RegExp(`\\b${wrong}\\b`, 'i').test(s.text));
    if (!hits.length) continue;
    out.push(makeFinding({
      rule: 'common-misspellings',
      title: `Spelling: "${wrong}" should be "${right}" (${hits.length} page${hits.length > 1 ? 's' : ''})`,
      severity: 'low',
      slides: hits.map((s) => s.index),
      why: 'Simple spelling errors cost nothing to fix and quietly reduce the professional impression of an otherwise strong portfolio.',
      fix: `Replace "${wrong}" with "${right}" on pages ${hits.map((s) => s.index).join(', ')}.`,
      evidence: `Found on pages ${hits.map((s) => s.index).join(', ')}.`,
      effort: 'minutes',
    }, {}));
  }
});

rule('title-too-long', (deck, out) => {
  for (const s of deck.slides) {
    if (!s.title) continue;
    const words = s.title.split(/\s+/).filter(Boolean).length;
    if (words <= THRESHOLDS.maxTitleWords && s.title.length <= 70) continue;
    out.push(makeFinding({
      rule: 'title-too-long',
      title: `Heading is too long to scan (${words} words)`,
      severity: 'low',
      why: 'A heading is a signpost. If it is a sentence, the page loses its signpost and the reader has to work out where they are.',
      fix: 'Shorten it to a few words, e.g. "Testing: drop test results" instead of a full sentence.',
      evidence: `"${s.title.slice(0, 100)}"`,
      effort: 'minutes',
    }, { slideNumber: s.index }));
  }
});

rule('deck-size', (deck, out) => {
  // Not a fault, but worth surfacing: page count strategy matters for a long submission.
  const slides = deck.slideCount;
  if (slides < 40) return;
  const emptyish = deck.slides.filter((s) => s.wordCount < THRESHOLDS.nearEmptyWords && s.imageCount === 0).length;
  if (emptyish < 5) return;
  out.push(makeFinding({
    rule: 'deck-size',
    title: `${emptyish} of ${slides} pages carry almost nothing`,
    severity: 'medium',
    slides: deck.slides.filter((s) => s.wordCount < THRESHOLDS.nearEmptyWords && s.imageCount === 0).map((s) => s.index),
    why: 'In a long portfolio, thin pages dilute the strong ones. A marker with 200 pages to read forms an impression from the average page, not the best one.',
    fix: 'Cut or merge the weakest pages. Consolidating 10 thin pages into 4 strong ones raises the overall standard of the submission with no new work.',
    evidence: `${emptyish} page(s) have fewer than ${THRESHOLDS.nearEmptyWords} words and no images.`,
    effort: 'hours',
  }, {}));
});

/* ------------------------------------------------------------------ *
 * Per-page change tracking
 *
 * A .pptx does not carry a reliable per-slide modification date: PowerPoint
 * rewrites the archive wholesale on save, so the ZIP timestamps and the
 * `p:cNvPr/@rev` counters do not tell you which PAGE changed. So the app keeps its
 * own record instead — which page had which fingerprint, and when that
 * fingerprint was last different.
 *
 * Fingerprints are compared by CONTENT, not by position. That matters: inserting
 * a page near the front shifts every later page's index, and a positional
 * comparison would report all of them as modified. Matching on the fingerprint
 * means an inserted page is "new" and the rest are correctly "unchanged".
 * ------------------------------------------------------------------ */

/**
 * Update a page timeline with the current scan's fingerprints.
 *
 * Pure, so it can be tested without a file. `history` is the previous state:
 * `{ '3': { fingerprint, firstSeen, lastChanged } }`. Returns the new history plus
 * a per-page status for the UI.
 *
 * Deleting a page loses its entry, because a page that has gone is not usefully
 * "last modified" any more.
 */
function updatePageHistory(history, slides, when) {
  const prev = history && typeof history === 'object' ? history : {};
  const now = when || new Date().toISOString();

  // Every fingerprint present last time, so an unchanged page is recognised
  // wherever it has moved to.
  const knownFingerprints = new Set();
  for (const key of Object.keys(prev)) {
    const entry = prev[key];
    if (entry && entry.fingerprint) knownFingerprints.add(entry.fingerprint);
  }

  const next = {};
  const pages = [];
  for (const slide of slides) {
    const index = String(slide.index);
    const fingerprint = slide.fingerprint;
    const before = prev[index];
    let status;
    let firstSeen;
    let lastChanged;

    if (before && before.fingerprint === fingerprint) {
      // Same content, same place.
      status = 'unchanged';
      firstSeen = before.firstSeen || now;
      lastChanged = before.lastChanged || now;
    } else if (knownFingerprints.has(fingerprint)) {
      /*
       * Same content, different place: the page was MOVED, or a page was inserted
       * or removed earlier in the deck. Its content has not changed, so it must
       * not be reported as modified — that is the whole point of matching on
       * content rather than position.
       */
      status = 'moved';
      const original = Object.keys(prev).map((k) => prev[k]).find((e) => e && e.fingerprint === fingerprint);
      firstSeen = (original && original.firstSeen) || now;
      lastChanged = (original && original.lastChanged) || now;
    } else if (before) {
      status = 'modified';
      firstSeen = before.firstSeen || now;
      lastChanged = now;
    } else {
      status = 'new';
      firstSeen = now;
      lastChanged = now;
    }

    next[index] = { fingerprint, firstSeen, lastChanged };
    pages.push({ number: slide.index, status, firstSeen, lastChanged });
  }

  return { history: next, pages };
}

/* ------------------------------------------------------------------ *
 * Runner
 * ------------------------------------------------------------------ */

/**
 * Run every rule and return consolidated findings.
 * Findings are merged by (rule, slide) so a rule firing twice on one page
 * becomes one action item, not two.
 */
/*
 * Rules that need a capability the source lacks, but whose inputs are exact even on a PDF, so they
 * can be trusted there.
 *
 * At module scope rather than inside runLocalChecks, because ruleCoverageFor has to answer the same
 * question before anything is run — and one definition is the only way the two cannot drift.
 */
const EXACT_ON_PDF = new Set(['text-too-small', 'wall-of-text', 'font-inconsistency']);

function runLocalChecks(deck, options = {}) {
  findingSeq = 0;
  const out = [];
  const errors = [];

  /*
   * A PDF has no speaker notes, so `notes` is never available for one.
   *
   * Geometry is more subtle. The extractor DOES recover real coordinates and real
   * font sizes from the PDF's drawing operators, so those rules become possible.
   * But the width of a line stays an ESTIMATE on the documents that matter most:
   * an export using Type0 CID fonts carries no /Widths table, so per-glyph
   * advances are not available without parsing the embedded font program.
   *
   * That splits the geometry rules in two, and the split is the whole point:
   *
   *   - Rules that read font SIZE and WORD COUNTS are working with exact values,
   *     so they run on a PDF normally.
   *   - Rules that read a WIDTH would be guessing, and a guess here produces
   *     either a flood of false "runs off the page" alarms or, if the guess is
   *     skewed the other way, silent misses. Measured against the real portfolio
   *     the estimate over-read a line by about 13%, which turned 106 correctly
   *     placed lines into off-page findings. So those rules stay SKIPPED and say
   *     why, rather than crying wolf about a page the student can see is fine.
   */
  const source = options.source || (deck && deck.source) || 'pptx';
  const capable = source === 'pdf'
    ? new Set(['text', 'images'])
    : new Set(['text', 'images', 'geometry', 'fonts', 'notes']);

  const skipped = [];
  /* Counted, not just skipped: a rule that is switched off must not appear in "checks that ran". */
  const disabledRuleCount = RULES.filter((r) => DISABLED_RULES.has(r.name)).length;

  for (const r of RULES) {
    if (DISABLED_RULES.has(r.name)) continue;
    const lacks = !capable.has(r.requires || 'text');
    if (lacks && !(source === 'pdf' && EXACT_ON_PDF.has(r.name))) {
      skipped.push({ rule: r.name, requires: r.requires });
      continue;
    }
    try {
      r.fn(deck, out, options);
    } catch (err) {
      errors.push({ rule: r.name, message: String(err && err.message ? err.message : err) });
    }
  }

  /*
   * Merge duplicates, keeping the highest severity.
   *
   * The key carries the TITLE as well as the rule and the page. Without it every deck-level finding
   * of a rule shared one key — "common-misspellings" files one finding per misspelling with no slide
   * number, so they all collided and all but the first were DELETED, with the survivor advertising
   * the union of every page number. The report named one typo, listed pages containing others, and
   * never mentioned them. Two findings are duplicates only if they say the same thing.
   */
  const merged = new Map();
  for (const f of out) {
    const key = `${f.rule}|${f.slide == null ? 'deck' : f.slide}|${f.title || ''}`;
    if (!merged.has(key)) {
      merged.set(key, f);
      continue;
    }
    const existing = merged.get(key);
    existing.slides = Array.from(new Set([...existing.slides, ...f.slides])).sort((a, b) => a - b);
    if (SEVERITY_ORDER[f.severity] < SEVERITY_ORDER[existing.severity]) existing.severity = f.severity;
    if (f.evidence && !existing.evidence) existing.evidence = f.evidence;
  }

  const findings = Array.from(merged.values()).sort((a, b) => {
    if (b.priority !== a.priority) return b.priority - a.priority;
    const as = a.slide == null ? 1e9 : a.slide;
    const bs = b.slide == null ? 1e9 : b.slide;
    return as - bs;
  });
  findings.forEach((f, i) => {
    f.rank = i + 1;
  });

  const bySeverity = { critical: 0, high: 0, medium: 0, low: 0, info: 0 };
  for (const f of findings) bySeverity[f.severity] = (bySeverity[f.severity] || 0) + 1;

  // Tier tally lets the UI separate "needs your thinking" from "quick tidies"
  // without re-deriving it, and lets it hide the obvious ones by default.
  const byTier = { substance: 0, presentation: 0, obvious: 0 };
  for (const f of findings) byTier[f.tier] = (byTier[f.tier] || 0) + 1;

  /*
   * A 0-100 number describing how much presentation tidying is outstanding.
   *
   * THIS IS NOT A MARK AND MUST NEVER BE SHOWN AS ONE. It was originally called
   * "score" and displayed as "n out of 100", which read as a grade. On a
   * portfolio that genuinely scored 99/100 it showed 51, which is indefensible:
   * these rules measure tidiness and completeness, not the quality of design
   * thinking, and they cannot know what a mark scheme will award.
   *
   * The UI now shows a plain count of points to check instead. The number is
   * retained (under an unambiguous name, plus a deprecated alias for older
   * saved reports) purely to colour the progress ring.
   */
  const pages = Math.max(1, deck.slideCount);
  const penalty = (bySeverity.critical * 12 + bySeverity.high * 4.5 + bySeverity.medium * 1.6 + bySeverity.low * 0.5) / pages * 10;
  const presentationTidiness = Math.max(0, Math.min(100, Math.round(100 - penalty)));

  return {
    findings,
    bySeverity,
    byTier,
    presentationTidiness,
    score: presentationTidiness, // deprecated alias; not a mark, do not display as one
    /*
     * The number of rules that actually RAN, not the number that exist. Reporting
     * the total on a PDF would claim checks that were skipped.
     *
     * Disabled rules are subtracted too. They are skipped by an early `continue` that never reaches
     * the skipped list, so they were silently counted as having run: the app claimed "all 18 free
     * checks" when two of them ('low-res-image', 'image-too-small') are switched off and only 16 can
     * ever execute. Telling a student their image resolution was checked when it was not is the
     * exact kind of claim this app must not make.
     */
    ruleCount: RULES.length - skipped.length - disabledRuleCount,
    totalRules: RULES.length - disabledRuleCount,
    skippedRules: skipped,
    source,
    disabledRules: Array.from(DISABLED_RULES),
    // Which board policy the checks ran under, so the report can say why the
    // caption rule was or was not applied.
    boardAnnotationRequired: options.annotationRequired !== false,
    // Per-page change tracking. Computed here because the fingerprints live here.
    pageHistory: updatePageHistory(options.pageHistory, deck.slides, options.now),
    errors,
    thresholds: THRESHOLDS,
  };
}

/**
 * Which rules WOULD run for a given kind of file, without a file to run them on.
 *
 * Everything that decides this is already above: the capability set is chosen from the SOURCE
 * alone, and the rules are fixed. Nothing here depends on the contents of a portfolio, which is
 * what makes it possible to say what a check will do BEFORE it is run — the thing the Overview
 * note is for.
 *
 * Kept next to runLocalChecks on purpose, and written to mirror it exactly: if the two ever
 * disagree, the note on the front page is describing a check that is not the one that runs.
 *
 * @param {string} source 'pdf' or 'pptx'
 */
function ruleCoverageFor(source) {
  const s = source === 'pdf' ? 'pdf' : 'pptx';
  const capable = s === 'pdf'
    ? new Set(['text', 'images'])
    : new Set(['text', 'images', 'geometry', 'fonts', 'notes']);
  const skipped = [];
  const disabledRuleCount = RULES.filter((r) => DISABLED_RULES.has(r.name)).length;
  for (const r of RULES) {
    if (DISABLED_RULES.has(r.name)) continue;
    const lacks = !capable.has(r.requires || 'text');
    if (lacks && !(s === 'pdf' && EXACT_ON_PDF.has(r.name))) {
      skipped.push({ rule: r.name, requires: r.requires });
    }
  }
  /*
   * Identical arithmetic to runLocalChecks, disabled rules included. There is a test that runs both
   * on the same deck and compares, because a front page that describes a different check from the
   * one that runs is worse than no front page.
   */
  return {
    source: s,
    ruleCount: RULES.length - skipped.length - disabledRuleCount,
    totalRules: RULES.length - disabledRuleCount,
    skippedRules: skipped,
  };
}

module.exports = {
  runLocalChecks,
  ruleCoverageFor,
  updatePageHistory,
  THRESHOLDS,
  estimateOverflow,
  RULE_TIER,
  DISABLED_RULES,
  TIERS,
};
