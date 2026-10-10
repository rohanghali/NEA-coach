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

  const line = fitLine(usable);
  const lo = Math.min(...usable.map((a) => a.app));
  const hi = Math.max(...usable.map((a) => a.app));

  /*
   * Outside the measured range, a fitted line is an extrapolation with no support.
   * Refuse to apply it and say so, rather than quietly inventing a number.
   */
  if (!line || raw < lo || raw > hi) {
    out.reason = 'outside-measured-range';
    out.note = `The app has only been checked against portfolios it scored between ${lo} and ${hi}, `
      + `and this one is outside that range, so no adjustment has been applied. `
      + `On the ones it has been checked against, this marking model read ${out.measuredLow} to `
      + `${out.measuredHigh} marks LOW. Treat this figure as a floor, not a ceiling.`;
    return out;
  }

  const adjusted = Math.round(line.slope * raw + line.intercept);
  // Never propose a mark above the paper's maximum, or below zero.
  out.calibrated = Math.max(0, Math.min(max, adjusted));
  out.applied = out.calibrated !== raw;
  out.reason = 'within-measured-range';
  /*
   * Worded without repeating the raw figure twice, which the previous version did:
   * "Adjusted from the AI's own 83 using ... which the raw figure read 2 to 6 low".
   */
  out.note = `The AI read this as ${raw}. On ${usable.length} portfolios whose real marks are known it read `
    + `${out.measuredLow} to ${out.measuredHigh} marks low, so it is shown higher here. `
    + `Those portfolios scored between ${lo} and ${hi} on the app.`;
  return out;
}

module.exports = { ANCHORS, fitLine, assessPrediction };
