'use strict';

/*
 * Turning a raw AI prediction into something the student can act on.
 *
 * THE MEASUREMENT
 *
 * The app has been checked against six real portfolios whose marks are known: two
 * marked 99 and 96 read by a cheap model, and three read by the default marking model
 * — 99, 96 and 89. What that showed:
 *
 *   - The cheap model (flash-lite) COMPRESSES. It gave a 99, a 96 and an 89 portfolio
 *     81, 80 and 66. A flat offset could never repair that, which is why the strong
 *     model is the default.
 *   - The default marking model (gemini-3.5-flash) is consistently LOW by 2-6 marks,
 *     and slightly over-confident: two of the three anchors fitted a line exactly,
 *     which proves nothing on its own, and the third at 89 is what tested it.
 *
 * WHAT THIS DOES ABOUT IT
 *
 * Three known points fitted a straight line to within 0.4 marks. That line is applied
 * here — but ONLY when the raw prediction falls inside the range the anchors cover.
 * Below that range nothing has been measured, so the raw figure is left alone and said
 * to be a guess. Extrapolating a correction into territory with no anchors would make
 * the app confidently wrong in a new way, which is worse than being visibly uncertain.
 *
 * Every number the student sees is labelled with where it came from, because a
 * predicted mark presented without its error bar invites more trust than it has earned.
 */

/**
 * The anchors: what the app said, and what was actually awarded.
 *
 * Kept here as data rather than buried in code, so adding a portfolio is a one-line
 * change and the reasoning stays visible. `tools/calibrate.js` holds the full records
 * including per-criterion marks.
 */
const ANCHORS = [
  // The default marking model, which runs 2-6 marks low.
  { app: 97, real: 99, label: 'AQA GCSE 99/100 portfolio', board: 'aqa-gcse', model: 'gemini-3.5-flash' },
  { app: 92, real: 96, label: 'AQA GCSE 96/100 portfolio', board: 'aqa-gcse', model: 'gemini-3.5-flash' },
  { app: 83, real: 89, label: 'AQA GCSE 89/100 portfolio', board: 'aqa-gcse', model: 'gemini-3.5-flash' },

  /*
   * The cheap model, on the SAME three portfolios. It does not merely read low, it
   * compresses: three portfolios really marked 99, 96 and 89 were scored 81, 80 and
   * 66. It needs a correction of a different shape, which is exactly why the anchors
   * are grouped by model — pooling them produced a nonsense fit
   * (real = app x 0.043 + 93.7).
   */
  { app: 81, real: 99, label: 'AQA GCSE 99/100 portfolio', board: 'aqa-gcse', model: 'gemini-3.1-flash-lite' },
  { app: 80, real: 96, label: 'AQA GCSE 96/100 portfolio', board: 'aqa-gcse', model: 'gemini-3.1-flash-lite' },
  { app: 66, real: 89, label: 'AQA GCSE 89/100 portfolio', board: 'aqa-gcse', model: 'gemini-3.1-flash-lite' },
];

/**
 * The marking models this app has actually been checked against, best first.
 *
 * A model with no anchors cannot be corrected at all, so its figure is the AI's raw
 * guess — measured at 2-6 marks low for the strong models and up to 23 marks low for
 * the cheap one. That difference is the whole reason this exists: when one marking
 * model will not answer and another has to be used instead, the replacement must be a
 * model the app can still correct.
 *
 * Ordered by the WORST error its fitted line leaves behind on real marked work, which
 * is the honest measure of how far to trust a corrected figure from it. Deriving the
 * order rather than writing it down means adding an anchor re-ranks the models on its
 * own.
 *
 * @returns {Array<{model: string, anchors: number, worstResidual: number}>}
 */
function calibratedModels() {
  const grouped = new Map();
  for (const a of ANCHORS) {
    if (!grouped.has(a.model)) grouped.set(a.model, []);
    grouped.get(a.model).push(a);
  }
  const out = [];
  for (const [model, list] of grouped) {
    // Two points are the least that can define a line.
    if (list.length < 2) continue;
    const fit = fitLine(list);
    if (!fit) continue;
    let worst = 0;
    for (const a of list) {
      const predicted = a.app * fit.slope + fit.intercept;
      worst = Math.max(worst, Math.abs(predicted - a.real));
    }
    out.push({ model, anchors: list.length, worstResidual: Math.round(worst * 10) / 10 });
  }
  out.sort((a, b) => a.worstResidual - b.worstResidual || a.model.localeCompare(b.model));
  return out;
}

/** Least-squares fit of real = app * slope + intercept. */
function fitLine(anchors) {
  const n = anchors.length;
  if (n < 2) return null;
  const meanX = anchors.reduce((s, a) => s + a.app, 0) / n;
  const meanY = anchors.reduce((s, a) => s + a.real, 0) / n;
  let num = 0;
  let den = 0;
  for (const a of anchors) {
    num += (a.app - meanX) * (a.real - meanY);
    den += (a.app - meanX) * (a.app - meanX);
  }
  // Every anchor at the same app mark means the slope is undetermined, not zero.
  if (den === 0) return null;
  const slope = num / den;
  return { slope, intercept: meanY - slope * meanX };
}

/**
 * Everything the UI needs to present a predicted mark honestly.
 *
 * @param {number} rawPredicted the app's own total
 * @param {number} maximum
 * @param {object} [opts] { board, model }
 */
function assessPrediction(rawPredicted, maximum, opts = {}) {
  const raw = Number(rawPredicted);
  const max = Number(maximum) || 100;
  const out = {
    raw: Number.isFinite(raw) ? raw : null,
    // Whether a correction was applied, and why or why not.
    calibrated: null,
    applied: false,
    reason: null,
    anchorsUsed: 0,
    measuredLow: null,
    measuredHigh: null,
    // How wide the measured error range is. A wide one means a correction would
    // be another guess, so none is applied.
    measuredSpan: null,
    // The plain sentence the UI shows under the number.
    note: null,
  };
  if (!Number.isFinite(raw)) return out;

  /*
   * Only the anchors for the board and marking model actually in use are evidence for
   * THIS prediction. Pooling across models is what produced a nonsense fit before
   * (real = app x 0.043 + 93.7), because different models compress differently.
   */
  const usable = ANCHORS.filter((a) => (!opts.board || a.board === opts.board)
    && (!opts.model || a.model === opts.model));
  out.anchorsUsed = usable.length;
  if (usable.length < 2) {
    out.reason = 'not-enough-anchors';
    out.note = 'No marked portfolios have been compared with this marking model yet, so this is the AI\'s own '
      + 'figure with no adjustment. Treat it as a rough guide.';
    return out;
  }

  const ups = usable.map((a) => a.real - a.app);
  out.measuredLow = Math.min(...ups);
  out.measuredHigh = Math.max(...ups);
  out.measuredSpan = out.measuredHigh - out.measuredLow;

  /*
   * A correction is only applied when the measured error is TIGHT.
   *
   * This is the rule that was missing, and its absence produced a figure the
   * student could see was nonsense: on one portfolio the app read 81, then
   * "corrected" it to 98 against a real mark of 90 - eight marks out, when the
   * raw reading had been only nine out and the adjustment was supposed to be an
   * improvement. Reported plainly as "still inaccurate ... this needs to go".
   *
   * The reason is arithmetic. The correction is only as precise as the spread of
   * the measured errors: if the model has been seen to read anywhere between 16
   * and 23 marks low, then adding a single number is another guess with a seven
   * mark spread. On the same portfolio two runs of the AI differed by 9 marks
   * BEFORE any correction, so a seven-mark-wide adjustment cannot be adding
   * information - it was adding error.
   *
   * The strong marking model's measured span is 4 marks, which is why it is
   * trustworthy; the cheap page-reader's is 7, which is why it is not.
   */
  const MAX_RELIABLE_SPAN = 5;
  if (out.measuredSpan > MAX_RELIABLE_SPAN) {
    out.reason = 'error-too-variable';
    out.note = `The AI read this as ${raw} out of ${max}. The app has measured this marking model as `
      + `reading anywhere between ${out.measuredLow} and ${out.measuredHigh} marks low, which is too wide `
      + `a range to correct for, so the AI's own figure is shown unchanged. Treat it as a rough guide.`;
    return out;
  }

  const line = fitLine(usable);
  const lo = Math.min(...usable.map((a) => a.app));
  const hi = Math.max(...usable.map((a) => a.app));

  /*
   * Outside the measured range, a fitted line is an extrapolation with no support.
   * Refuse to apply it and say so, rather than quietly inventing a number.
   */
  if (!line || raw < lo || raw > hi) {
    out.reason = 'outside-measured-range';
    /*
     * "it scored between 83 and 97" is the APP's reading, and saying it as "the app
     * has only been checked against portfolios it scored between..." left the reader
     * to guess whether those were the real marks or the app's. They are the app's.
     * The real marks are the second range. Both are now labelled.
     */
    out.note = `The app has only been checked against portfolios it scored between ${lo} and ${hi} out of ${max}, and this one came out at ${raw}, `
      + `so no adjustment has been applied. On those known portfolios, `
      + `${opts && opts.model ? 'the model that marked this (' + opts.model + ')' : 'this marking model'} `
      + `read ${out.measuredLow} to ${out.measuredHigh} marks low. Treat this figure as a floor, not a ceiling.`;
    return out;
  }

  const adjusted = Math.round(line.slope * raw + line.intercept);
  // Never propose a mark above the paper's maximum, or below zero.
  out.calibrated = Math.max(0, Math.min(max, adjusted));
  out.applied = out.calibrated !== raw;
  out.reason = 'within-measured-range';
  /*
   * The sentence used to say "it read 16 to 23 marks low ... Those portfolios
   * scored between 66 and 81 on the app", which mixes three different quantities in
   * one breath: the AI's reading in marks, its error in marks, and the anchor scores
   * in PERCENT. A reader cannot tell which number refers to what, and that is how
   * "it read 16 to 23 marks low" ended up printed above a table totalling 75.
   *
   * Everything is now in one unit - marks out of the paper total - and the two
   * figures that matter are named: what the AI said, and what the anchors say it
   * usually misses by.
   */
  out.note = `The AI read this as ${raw} out of ${max}. Measured against `
    + `${usable.length} portfolios whose real marks are known, `
    + `${opts && opts.model ? 'the model that marked this (' + opts.model + ')' : 'this marking model'} `
    + `reads ${out.measuredLow} to ${out.measuredHigh} marks low on a paper worth ${max}, so the `
    + `figure is adjusted up to ${out.calibrated}.`;
  return out;
}

module.exports = { ANCHORS, fitLine, assessPrediction, calibratedModels };
