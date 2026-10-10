#!/usr/bin/env node
'use strict';
/*
 * tools/calibrate.js — measure how accurate the predicted mark actually is.
 *
 * NEA Coach predicts a mark. This tool answers the only question that matters
 * about that prediction: is it CLOSE TO THE TRUTH? To do that it needs an
 * anchor — a portfolio whose per-criterion marks are known, from a teacher or a
 * moderator — and it compares the app's judgement against those real marks.
 *
 * This exists because the app cannot calibrate itself. It can prove it is
 * self-consistent (repeat runs agree), and it can prove it is not wildly wrong,
 * but "accurate" is meaningless without something known to be right. Every
 * attempt to tune a prompt without an anchor is guesswork dressed as progress:
 * it changes the number, and nobody can tell whether it changed it towards the
 * truth.
 *
 * USAGE
 *   node tools/calibrate.js run <pdf> --marks a=10,b=10,c=20,d=20,e=20,f=19 \
 *        [--board aqa-gcse] [--label "GCSE 99/100"] [--runs 3] [--model NAME]
 *   node tools/calibrate.js run <pdf> --total 99 [--of 100] [--label "..."]
 *   node tools/calibrate.js list
 *   node tools/calibrate.js show <id>
 *
 * Two kinds of anchor, and they are NOT equal:
 *
 *   --marks  the per-criterion marks. This is the real measurement: it shows
 *            WHICH criteria are wrong, so it can distinguish a uniform offset
 *            from an uneven error, and those need different fixes.
 *   --total  only the overall mark. Still useful — it measures the size of the
 *            error and the direction — but it CANNOT say where the error is. A
 *            single total cannot tell a uniform 18-mark shortfall apart from a
 *            shortfall concentrated in one criterion. Everything reported from a
 *            total-only anchor is labelled so it cannot be mistaken for the more
 *            precise comparison.
 *
 * COST: one run costs roughly 1.2p on gemini-3.1-flash-lite for a 60-page deck,
 * so the default three runs is a few pence. The tool prints the measured cost
 * afterwards, taken from the API's own token counts rather than estimated.
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const ai = require(path.join(ROOT, 'lib', 'ai.js'));
const STATE_DIR = path.join(ROOT, 'state');
const CALIBRATION_PATH = path.join(STATE_DIR, 'calibration.json');

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

function fail(message) {
  console.error('\n  ' + message + '\n');
  process.exit(1);
}

function argValue(args, name) {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] && !args[i + 1].startsWith('--') ? args[i + 1] : null;
}

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (err) {
    return fallback;
  }
}

function writeJson(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2));
  fs.renameSync(tmp, file);
}

/** "a=10,b=10,c=20" -> { a: 10, b: 10, c: 20 } */
function parseMarks(text) {
  if (!text) return null;
  const out = {};
  for (const part of String(text).split(',')) {
    const m = part.match(/^\s*([A-Za-z0-9]+)\s*[=:]\s*(-?\d+(?:\.\d+)?)\s*$/);
    if (!m) return null;
    out[m[1]] = Number(m[2]);
  }
  return Object.keys(out).length ? out : null;
}

function pct(n) {
  return n == null ? '—' : (n >= 0 ? '+' : '') + Math.round(n * 10) / 10;
}

/* ------------------------------------------------------------------ *
 * The comparison
 * ------------------------------------------------------------------ */

/**
 * Compare predicted marks against known marks.
 *
 * Kept pure and exported so the tests can check the arithmetic without calling
 * the API. It reports, per criterion, the difference between what the app said
 * and what was actually awarded, and then summarises the pattern:
 *
 *   bias     — the mean signed difference. A consistent bias is the most
 *              fixable kind of error: the whole scale is shifted.
 *   spread   — how much the differences vary. Large spread means the app is
 *              unreliable per criterion, which averaging cannot fix.
 *   worst    — the criterion it is most wrong about, by absolute difference.
 *
 * A "marks" comparison only makes sense where the real marks are on the same
 * scale, so both sides are also expressed as a percentage of the maximum.
 */
function compareToKnown(criteria, predictedRows, known) {
  const rows = [];
  for (const crit of criteria || []) {
    const id = String(crit.id);
    if (!(id in known)) continue;
    const max = Number(crit.weight) || null;
    const predicted = (predictedRows || []).find((r) => String(r.id) === id);
    const predictedMark = predicted && typeof predicted.mark === 'number' ? predicted.mark : null;
    const actual = Number(known[id]);
    const diff = predictedMark == null ? null : predictedMark - actual;
    rows.push({
      id,
      name: crit.name,
      max,
      predicted: predictedMark,
      actual,
      diff,
      diffPercent: (diff == null || !max) ? null : (diff / max) * 100,
      predictedPercent: (predictedMark == null || !max) ? null : (predictedMark / max) * 100,
      actualPercent: (!max) ? null : (actual / max) * 100,
    });
  }

  const withDiff = rows.filter((r) => r.diff != null);
  const diffs = withDiff.map((r) => r.diff);
  /*
   * SIGNED mean and the standard deviation around it.
   *
   * The signed mean matters: a mean of 0 achieved by being 8 high on one criterion
   * and 8 low on another is not "no bias", it is two errors. The standard
   * deviation is what separates a uniform offset (near 0) from an uneven one.
   * `meanAbsBias` is reported alongside for size, and is never used to classify —
   * using the mean ABSOLUTE difference as the "bias" made its own spread come out
   * at roughly |mean|, which called a perfectly uniform 3-mark offset "uneven".
   */
  const mean = diffs.length ? diffs.reduce((a, b) => a + b, 0) / diffs.length : null;
  const spread = diffs.length > 1
    ? Math.sqrt(diffs.reduce((s, d) => s + (d - mean) ** 2, 0) / diffs.length)
    : 0;

  const totalPredicted = withDiff.reduce((s, r) => s + (r.predicted || 0), 0);
  const totalActual = withDiff.reduce((s, r) => s + r.actual, 0);
  const totalMax = withDiff.reduce((s, r) => s + (r.max || 0), 0);
  const worst = withDiff.slice().sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff))[0] || null;
  const meanAbs = diffs.length ? diffs.reduce((s, d) => s + Math.abs(d), 0) / diffs.length : null;

  /*
   * Classify the ERROR PATTERN, because the patterns need different responses and
   * a bare "it's wrong" is not actionable:
   *
   *   well calibrated   — nothing to fix.
   *   consistent offset — the whole scale is shifted by roughly the same amount
   *                       everywhere, which recalibration can correct. Needs a
   *                       SMALL spread: a large average error means nothing if it
   *                       is achieved by being high on some criteria and low on
   *                       others, which averages out but is not a fixable offset.
   *   uneven error      — wrong by different amounts in different places. An offset
   *                       correction would make some criteria worse, so the fix has
   *                       to be per criterion. See `worst` and `rows`.
   *   unreliable        — large in size AND highly variable; no simple correction.
   *
   * Both magnitude and variation are required for the "consistent" label. Checking
   * only one of them produces absurd results in the other direction: an average of
   * near zero with a huge spread looks "consistent" to a magnitude-only rule, and
   * a uniform 3-mark offset looks "uneven" to a spread-only rule. Both mistakes
   * were made here and are covered by tests.
   *
   * Deliberately NOT using an outlier rule to detect "one bad criterion": on the
   * one portfolio measured so far that rule fired on a spread that was genuinely
   * broad, and would have sent us chasing a single criterion when the error was
   * spread across several. The per-criterion rows are printed in full so a human
   * can see the shape without the tool guessing at it.
   */
  /*
   * `meanAbs` is the typical size of the error; `spread` is how unevenly it is
   * distributed. Both matter, and they answer different questions:
   *   - An error that is EVEN in shape but large in size is still correctable by
   *     recalibration, but calling a 5-mark typical error "recalibratable" invites
   *     the student to trust a mark that is nowhere near right, so size gets its
   *     own band.
   *   - An error that is uneven cannot be corrected by any single adjustment,
   *     however small its average.
   */
  let verdict;
  if (mean == null) verdict = 'no comparable criteria';
  else if (Math.abs(mean) < 1.2 && spread < 2) verdict = 'well calibrated';
  else if (spread < 2 && meanAbs != null && meanAbs < 3.5) verdict = 'consistent offset — recalibratable';
  else if (spread < 4) verdict = 'uneven error — check the worst criteria';
  else verdict = 'erratic error — no simple correction applies';

  return {
    rows,
    meanBias: mean,
    meanAbsBias: meanAbs,
    biasSpread: spread,
    totalPredicted,
    totalActual,
    totalMax,
    totalDiff: totalPredicted - totalActual,
    predictedPercent: totalMax ? (totalPredicted / totalMax) * 100 : null,
    actualPercent: totalMax ? (totalActual / totalMax) * 100 : null,
    worst: worst ? { id: worst.id, name: worst.name, diff: worst.diff } : null,
    verdict,
  };
}

/**
 * Compare against a KNOWN TOTAL only.
 *
 * Weaker than a per-criterion comparison and must never be presented as if it
 * were the same thing: it can measure the size and direction of the error but not
 * its distribution, so it cannot tell a uniform shortfall from one concentrated
 * in a single criterion. What it CAN do is establish the scale correction, which
 * is the first thing worth knowing — and it is all that is available when a
 * teacher's total is known but the mark sheet is not.
 */
function compareToKnownTotal(criteria, predictedRows, total, maxTotal) {
  const rows = (predictedRows || []).filter((r) => typeof r.mark === 'number');
  const totalPredicted = rows.reduce((s, r) => s + r.mark, 0);
  const availableFromCriteria = (criteria || []).reduce((s, c) => s + (Number(c.weight) || 0), 0);
  // Prefer the mark scheme's own maximum; fall back to the stated one.
  const totalMax = maxTotal || availableFromCriteria || null;
  const scale = totalMax && availableFromCriteria && totalMax !== availableFromCriteria
    ? totalMax / availableFromCriteria
    : 1;
  const predictedOnSameScale = totalPredicted * scale;

  return {
    mode: 'total-only',
    rows,
    totalPredicted: Math.round(predictedOnSameScale * 10) / 10,
    totalActual: total,
    totalMax,
    totalDiff: Math.round((predictedOnSameScale - total) * 10) / 10,
    predictedPercent: totalMax ? (predictedOnSameScale / totalMax) * 100 : null,
    actualPercent: totalMax ? (total / totalMax) * 100 : null,
    // The correction implied by this anchor: add this many marks to the app's total.
    impliedOffset: Math.round((total - predictedOnSameScale) * 10) / 10,
    meanBias: null,
    meanAbsBias: null,
    biasSpread: null,
    worst: null,
    /*
     * Plain words, sign included, because "app is -19 marks off" is not something
     * a student should have to decode.
     */
    verdict: `total-only anchor: the app is ${Math.abs(Math.round((predictedOnSameScale - total) * 10) / 10)} marks `
      + `${predictedOnSameScale < total ? 'LOW' : 'HIGH'} on this portfolio. It cannot say which criteria are `
      + 'responsible — a per-criterion split is needed for that.'
      + (scale !== 1 ? ` (totals rescaled from ${availableFromCriteria} to ${totalMax})` : ''),
  };
}

/* ------------------------------------------------------------------ *
 * Fitting a correction from several anchors
 * ------------------------------------------------------------------ */

/**
 * Derive the correction from pairs of (app total, real total).
 *
 * ONE anchor can only tell you the offset, and assuming that offset holds
 * everywhere is a guess: a portfolio the app rates 20 might really be 40, or
 * really be 25, and one point cannot distinguish those. TWO anchors at different
 * marks are the minimum that can tell a constant offset from a proportional
 * error, which are fixed in completely different ways:
 *
 *   constant offset    real = app + c      (add the same marks everywhere)
 *   proportional       real = app * k      (the error grows with quality)
 *
 * This fits the straight line `real = app * k + c` through the anchors and reports
 * both the fit and how well it holds, because with two points a line always fits
 * perfectly — the residuals only mean something from three anchors onwards. If
 * the fit says the error is proportional, a constant offset would OVER-correct
 * weak portfolios and UNDER-correct strong ones, so this distinction matters.
 *
 * @returns {{k:number, c:number, mode:string, residuals:Array, maxResidual:number, n:number}}
 */
function fitCorrection(pairs) {
  const pts = (pairs || []).filter((p) => Number.isFinite(p.app) && Number.isFinite(p.real));
  if (!pts.length) return null;
  if (pts.length === 1) {
    const c = pts[0].real - pts[0].app;
    return {
      k: 1, c, mode: Math.abs(c) < 1 ? 'appears well calibrated' : 'constant offset',
      residuals: [{ app: pts[0].app, real: pts[0].real, predicted: pts[0].real, residual: 0 }],
      maxResidual: 0, n: 1,
      caveat: 'One anchor only: this assumes the error is the same for every portfolio, which one point cannot prove.',
    };
  }

  const n = pts.length;
  const meanX = pts.reduce((s, p) => s + p.app, 0) / n;
  const meanY = pts.reduce((s, p) => s + p.real, 0) / n;
  const variance = pts.reduce((s, p) => s + (p.app - meanX) ** 2, 0);

  // Two anchors at different marks give a line. Two at the SAME mark give only an
  // offset, because the slope is not determined.
  const k = variance === 0 ? 1 : pts.reduce((s, p) => s + (p.app - meanX) * (p.real - meanY), 0) / variance;
  const c = meanY - k * meanX;

  const residuals = pts.map((p) => {
    const predicted = p.app * k + c;
    return {
      app: p.app, real: p.real,
      predicted: Math.round(predicted * 10) / 10,
      residual: Math.round((p.real - predicted) * 10) / 10,
    };
  });
  const maxResidual = residuals.reduce((m, r) => Math.max(m, Math.abs(r.residual)), 0);

  let mode;
  if (maxResidual < 1.5) mode = 'the line fits both anchors, so the correction is sound for portfolios in this range';
  else mode = 'the anchors disagree with a straight line, so no simple correction is safe';

  /*
   * How far apart the anchors sit on the app's own scale.
   *
   * This is what makes a slope trustworthy or absurd. Two anchors at app 81 and 80
   * for real 99 and 96 imply THREE real marks per app mark — a cliff, not a
   * calibration — purely because the app happened to rate two very different
   * portfolios almost identically. A flat offset would be the honest reading of
   * that data, and the tool must say so rather than hand over a wild slope.
   */
  const appRange = Math.max(...pts.map((p) => p.app)) - Math.min(...pts.map((p) => p.app));
  const narrow = n >= 2 && appRange < 5;

  if (narrow) mode = 'the anchors are too close together on the app\'s scale to support a slope — treat the error as a flat offset';
  else if (maxResidual >= 1.5) mode = 'no straight line fits all the anchors, so no simple correction is safe';

  return {
    k: Math.round(k * 1000) / 1000,
    c: Math.round(c * 10) / 10,
    mode,
    residuals,
    maxResidual: Math.round(maxResidual * 10) / 10,
    n,
    appRange: Math.round(appRange * 10) / 10,
    /* The flat offset these anchors imply, which is the safer reading when narrow. */
    flatOffset: Math.round((meanY - meanX) * 10) / 10,
    narrow,
    caveat: n === 2
      ? (narrow
        ? 'Two anchors fit a line exactly, and these two sit only ' + Math.round(appRange * 10) / 10
          + ' marks apart on the app\'s scale, so the slope is an artefact. A third portfolio that the app rates much lower or higher is what would fix it.'
        : 'Two anchors fit a line exactly, so the residuals prove nothing yet. A third portfolio at a different mark would test it.')
      : null,
  };
}

/** Format a fitted correction as an instruction a human can apply. */
function describeCorrection(fit) {
  if (!fit) return 'no anchors recorded';
  if (fit.n === 1) {
    return `add ${fit.c >= 0 ? '+' : ''}${fit.c} marks to the app's total (one anchor, so this is an assumption)`;
  }
  /*
   * When the anchors are bunched together, the slope is an artefact of that and
   * quoting it would be actively misleading. Lead with the flat offset instead.
   */
  if (fit.narrow) {
    return `add about ${fit.flatOffset >= 0 ? '+' : '−'}${Math.abs(fit.flatOffset)} marks to the app's total — `
      + `the fitted slope (×${fit.k}) comes from anchors only ${fit.appRange} marks apart, so it is an artefact and not a real relationship`;
  }
  const parts = [];
  parts.push(`real ≈ app × ${fit.k} ${fit.c >= 0 ? '+' : '−'} ${Math.abs(fit.c)}`);
  if (Math.abs(fit.k - 1) < 0.02) {
    parts.push(`which is close to a constant ${fit.c >= 0 ? '+' : '−'}${Math.abs(Math.round(fit.c))} marks, so a flat offset would do`);
  } else if (Math.abs(fit.c) < 3) {
    parts.push('which is close to proportional, so a flat offset would OVER-correct weaker portfolios');
  } else {
    parts.push('which is neither purely flat nor purely proportional');
  }
  return parts.join(' — ');
}

/* ------------------------------------------------------------------ *
 * Running a calibration
 * ------------------------------------------------------------------ */

/**
 * Build one calibration record from a visual review's coverage.
 *
 * Split out of runCalibration so the whole comparison path — predicting marks,
 * matching them to known marks, judging the pattern, pricing it — can be tested
 * WITHOUT spending money on an API call. Anything that only runs after a paid
 * call is code that gets written once and never verified.
 */
function buildRecord({ label, board, model, runs, pdfPath, pdfBytes, known, knownTotal, totalMax, spec, vision, seconds }) {
  const predicted = ai.summarisePredictedMarks(spec.criteria, vision.coverage);
  const usingTotalOnly = !known && knownTotal != null;
  const comparison = usingTotalOnly
    ? compareToKnownTotal(spec.criteria, predicted.rows, knownTotal, totalMax)
    : compareToKnown(spec.criteria, predicted.rows, known || {});
  const cost = ai.costFor(model, vision.tokensIn || 0, vision.tokensOut || 0);

  /*
   * The real total is only comparable over the criteria that were actually
   * judged, so it is summed from the same rows the comparison used rather than
   * from everything in the mark scheme.
   */
  const judged = comparison.rows.filter((r) => known ? r.diff != null : typeof r.mark === 'number');
  const availableFromCriteria = judged.reduce((s, r) => s + (r.max || r.weight || 0), 0);
  const finalMax = comparison.totalMax || availableFromCriteria || null;

  return {
    id: 'cal-' + Date.now(),
    at: new Date().toISOString(),
    label,
    board,
    model,
    runs: vision.runs || runs,
    pdf: pdfPath,
    pdfMB: pdfBytes ? Number((pdfBytes / 1048576).toFixed(1)) : null,
    // Which kind of anchor this is. Never let a total-only result be read as the
    // more precise per-criterion comparison.
    anchor: usingTotalOnly ? 'total' : 'per-criterion',
    known: known || null,
    knownTotal: knownTotal != null ? knownTotal : null,
    predicted: predicted.rows.map((r) => ({ id: r.id, mark: r.mark, max: r.max })),
    totalPredicted: comparison.totalPredicted,
    totalMaximum: finalMax,
    totalActual: comparison.totalActual,
    comparison,
    tokensIn: vision.tokensIn || 0,
    tokensOut: vision.tokensOut || 0,
    costGbp: cost.gbp,
    seconds,
  };
}

/**
 * Record an anchor from a measurement made OUTSIDE this tool.
 *
 * The `run` command does the measuring, but a measurement is often taken by hand
 * (or by a one-off script) before anyone thinks to log it. Hand-building the
 * record afterwards is how a mean got written where the app reports a median, and
 * how an anchor ended up recorded with no predicted marks at all — so there is a
 * command for it instead.
 *
 *   node tools/calibrate.js record --app 95 --real 89 --marks a=8,b=9,... \
 *        [--board aqa-gcse] [--label "..."] [--model gemini-3.5-flash] [--runs 96,95,95]
 */
function recordAnchor(args) {
  const appTotal = Number(argValue(args, '--app'));
  const realTotal = Number(argValue(args, '--real'));
  if (!Number.isFinite(appTotal) || !Number.isFinite(realTotal)) {
    fail('Need --app <the app\'s total> and --real <the mark actually awarded>.'
      + '\n  --marks is optional but strongly preferred: without a per-criterion split'
      + '\n  the anchor can only show the SIZE of the error, not where it is.');
  }
  const board = argValue(args, '--board') || readJson(path.join(STATE_DIR, 'config.json'), {}).specId || 'eduqas-alevel';
  const spec = ai.loadSpec(path.join(ROOT, 'spec'), board);
  const marks = parseMarks(argValue(args, '--marks'));
  const runsRaw = argValue(args, '--runs');
  const runs = runsRaw ? String(runsRaw).split(',').map(Number).filter(Number.isFinite) : null;

  // If per-criterion marks are given, they must cover the scheme and sum to --app.
  if (marks) {
    const problem = checkMarksConsistent(spec.criteria, marks, appTotal);
    if (problem) fail(problem);
  }

  const predicted = marks
    ? spec.criteria.map((c) => ({ id: String(c.id), mark: Number(marks[c.id]), max: c.weight }))
    : null;
  const comparison = marks
    ? compareToKnown(spec.criteria, predicted, {})
    : compareToKnownTotal(spec.criteria, [], realTotal, 100);

  const rec = {
    id: 'cal-' + Date.now(),
    at: new Date().toISOString(),
    label: argValue(args, '--label') || 'recorded anchor',
    board,
    modelUsedForMarking: argValue(args, '--model') || null,
    runs: runs ? runs.length : null,
    anchor: marks ? 'per-criterion' : 'total',
    known: null,
    knownTotal: realTotal,
    predicted,
    totalPredicted: Math.round(appTotal),
    totalMaximum: 100,
    totalActual: realTotal,
    observedTotals: runs,
    comparison,
    note: argValue(args, '--note') || 'Recorded from a measurement taken outside this tool.',
    costGbp: null, tokensIn: null, tokensOut: null,
  };

  const store = readJson(CALIBRATION_PATH, { runs: [] });
  store.runs = Array.isArray(store.runs) ? store.runs : [];
  store.runs.push(rec);
  writeJson(CALIBRATION_PATH, store);
  console.log('');
  console.log('  Recorded ' + rec.id + ': app ' + rec.totalPredicted + ' vs real ' + rec.totalActual);
  if (!marks) {
    console.log('  No --marks given, so this anchor shows the size of the error but not where it is.');
  }
  console.log('');
}

/**
 * Is a set of per-criterion marks usable as an anchor?
 *
 * Returns a message describing the problem, or null when it is fine. Split out so
 * it can be tested: an anchor whose parts do not add up to its total silently
 * corrupts every fit built on it, and that is not obvious from the output.
 */
function checkMarksConsistent(criteria, marks, appTotal) {
  const missing = (criteria || []).map((c) => String(c.id)).filter((id) => !(id in marks));
  if (missing.length) return 'No mark given for: ' + missing.join(', ') + '. Every criterion needs one.';
  const sum = (criteria || []).reduce((t, c) => t + Number(marks[String(c.id)] || 0), 0);
  if (Number.isFinite(appTotal) && sum !== Math.round(appTotal)) {
    return 'The per-criterion marks add up to ' + sum + ' but --app says ' + appTotal
      + '. One of them is wrong; recording both would corrupt the fit.';
  }
  // A mark above the criterion's own maximum is always a mistake.
  for (const c of criteria || []) {
    const v = Number(marks[String(c.id)]);
    if (Number.isFinite(v) && v > Number(c.weight)) {
      return 'Criterion ' + c.id + ' is given ' + v + ' but is only out of ' + c.weight + '.';
    }
  }
  return null;
}

async function runCalibration(args) {
  const pdfPath = args[0];
  if (!pdfPath) fail('Give me the portfolio PDF: node tools/calibrate.js run <pdf> --marks a=10,...');
  if (!fs.existsSync(pdfPath)) fail('No such file: ' + pdfPath);

  const known = parseMarks(argValue(args, '--marks'));
  const knownTotalRaw = argValue(args, '--total');
  const knownTotal = knownTotalRaw != null ? Number(knownTotalRaw) : null;

  if (!known && knownTotal == null) {
    fail('Give me the real marks, either per criterion or as a total.\n'
      + '  per criterion: --marks a=10,b=10,c=20,d=20,e=20,f=19\n'
      + '  total only   : --total 99 --of 100');
  }
  if (knownTotal != null && !Number.isFinite(knownTotal)) fail('--total must be a number, e.g. --total 99');
  if (known && knownTotal != null) {
    fail('Give either --marks or --total, not both. The per-criterion marks are more useful.');
  }

  const cfg = readJson(path.join(STATE_DIR, 'config.json'), {});
  const apiKey = cfg.apiKey;
  if (!apiKey) fail('No API key in state/config.json, so nothing can be judged.');

  const board = argValue(args, '--board') || cfg.specId || 'eduqas-alevel';
  const label = argValue(args, '--label') || path.basename(pdfPath);
  const runs = Number(argValue(args, '--runs')) || 3;
  const model = argValue(args, '--model') || cfg.model;
  const ofTotal = argValue(args, '--of') != null ? Number(argValue(args, '--of')) : null;

  const spec = ai.loadSpec(path.join(ROOT, 'spec'), board);
  if (!spec.criteria.length) fail('No criteria loaded for board "' + board + '".');

  if (known) {
    const missing = spec.criteria.map((c) => String(c.id)).filter((id) => !(id in known));
    if (missing.length) {
      fail('You did not give a mark for: ' + missing.join(', ')
        + '\n  Every criterion needs a real mark or the comparison is incomplete.\n'
        + '  If you only know the overall mark, use --total instead.\n'
        + '  Criteria for this board: ' + spec.criteria.map((c) => c.id + ' (out of ' + c.weight + ')').join(', '));
    }
  }

  const pdf = fs.readFileSync(pdfPath);
  console.log('');
  console.log('  Calibrating: ' + label);
  console.log('  board      : ' + spec.label);
  console.log('  runs       : ' + runs + ' on ' + model);
  console.log('  anchor     : ' + (known
    ? 'per-criterion (' + spec.criteria.map((c) => c.id + '=' + known[c.id]).join(', ') + ')'
    : 'TOTAL ONLY (' + knownTotal + (ofTotal ? ' of ' + ofTotal : '') + ') — cannot show WHICH criteria'));
  console.log('');

  /*
   * A synthetic deck is enough: the visual review only needs the page count for
   * bounds-checking the page numbers the model returns.
   */
  const deck = { slideCount: 60, slides: [] };
  const started = Date.now();
  const vision = await ai.runVisionPass({
    deck,
    apiKey,
    model,
    spec,
    pdf,
    pdfName: path.basename(pdfPath),
    runs,
    onLog: (m) => console.log('  ' + m),
  });
  if (!vision || !vision.ok) fail('The visual review did not return anything usable.');

  const record = buildRecord({
    label, board, model, runs, pdfPath, pdfBytes: pdf.length, known,
    knownTotal, totalMax: ofTotal, spec, vision,
    seconds: Math.round((Date.now() - started) / 1000),
  });

  const store = readJson(CALIBRATION_PATH, { runs: [] });
  store.runs = Array.isArray(store.runs) ? store.runs : [];
  store.runs.push(record);
  writeJson(CALIBRATION_PATH, store);

  printComparison(record);
  console.log('  Saved as ' + record.id + ' in state/calibration.json.');
  console.log('');
}

function printComparison(record) {
  const c = record.comparison;
  const totalOnly = record.anchor === 'total';
  console.log('');
  console.log('  ' + record.label + '  (' + record.board + ', ' + record.runs + ' run(s), '
    + (totalOnly ? 'TOTAL-ONLY anchor' : 'per-criterion anchor') + ')');
  console.log('  ' + '-'.repeat(66));

  if (!totalOnly) {
    console.log('  ' + 'criterion'.padEnd(34) + 'app   real   diff');
    for (const r of c.rows) {
      const name = String(r.name || r.id).slice(0, 33);
      console.log('  ' + name.padEnd(34)
        + String(r.predicted == null ? '—' : r.predicted).padStart(4)
        + String(r.actual).padStart(7)
        + String(pct(r.diff)).padStart(7));
    }
    console.log('  ' + '-'.repeat(66));
  }

  console.log('  ' + 'TOTAL'.padEnd(34)
    + String(c.totalPredicted).padStart(4)
    + String(c.totalActual).padStart(7)
    + String(pct(c.totalDiff)).padStart(7));
  console.log('');
  console.log('  predicted : ' + Math.round(c.predictedPercent) + '%   real: ' + Math.round(c.actualPercent) + '%'
    + '   (' + record.totalActual + ' known, out of ' + record.totalMaximum + ')');

  if (totalOnly) {
    console.log('  offset    : ' + pct(c.impliedOffset) + ' marks (what would have to be added to the app\'s total)');
    console.log('');
    console.log('  NOTE: a total-only anchor measures the SIZE of the error, not where it');
    console.log('  is. An 18-mark shortfall spread evenly across six criteria and one');
    console.log('  concentrated in two criteria look identical here. For that, the');
    console.log('  per-criterion marks are needed.');
  } else {
    console.log('  mean bias : ' + pct(c.meanBias) + ' marks per criterion');
    console.log('  spread    : ' + (Math.round(c.biasSpread * 10) / 10) + ' marks');
    if (c.worst) console.log('  worst     : ' + c.worst.name + ' (' + pct(c.worst.diff) + ')');
  }
  console.log('  VERDICT   : ' + c.verdict);
  console.log('');
  /*
   * Only print the cost line when the record actually carries a measurement. A
   * hand-entered baseline has no token counts, and printing "undefineds" makes a
   * good record look broken.
   */
  const bits = [];
  if (record.costGbp != null) bits.push('£' + record.costGbp.toFixed(4) + ' (' + (record.costGbp * 100).toFixed(2) + 'p)');
  if (record.seconds != null) bits.push(record.seconds + 's');
  if (record.tokensIn != null) bits.push(record.tokensIn + ' in / ' + record.tokensOut + ' out');
  if (bits.length) console.log('  cost      : ' + bits.join(' | '));
  console.log('');
}

/* ------------------------------------------------------------------ *
 * Listing and showing
 * ------------------------------------------------------------------ */

function listRuns() {
  const store = readJson(CALIBRATION_PATH, { runs: [] });
  const runs = store.runs || [];
  if (!runs.length) {
    console.log('\n  No calibrations recorded yet.\n');
    console.log('  Run one with:');
    console.log('    node tools/calibrate.js run <pdf> --marks a=10,b=10,c=20,d=20,e=20,f=19\n');
    return;
  }
  console.log('');
  console.log('  ' + 'id'.padEnd(18) + 'when'.padEnd(18) + 'board'.padEnd(14) + 'app / real');
  for (const r of runs) {
    console.log('  ' + String(r.id).padEnd(18)
      + String((r.at || '').slice(0, 16)).padEnd(18)
      + String(r.board).padEnd(14)
      + String(r.totalPredicted) + ' / ' + String(r.totalActual)
      + '   [' + String(r.anchor || 'per-criterion') + '] '
      + (r.comparison ? r.comparison.verdict : ''));
  }
  console.log('');
}

/**
 * Fit a correction across every anchor that has both a predicted and a real total.
 *
 * Deliberately refuses to produce a line from anchors of different quarters: a
 * GCSE mark and an A-level mark are not the same scale and fitting across them
 * would invent a relationship. Only records whose board matches are used.
 */
function fitAll() {
  const store = readJson(CALIBRATION_PATH, { runs: [] });
  const runs = (store.runs || []).filter((r) => Number.isFinite(r.totalPredicted) && Number.isFinite(r.totalActual));
  if (!runs.length) {
    console.log('\n  No anchors with both a predicted and a real total.\n');
    return;
  }
  /*
   * Group by board AND by the model that did the marking.
   *
   * Records made with different marking models must never be fitted together:
   * gemini-3.1-flash-lite was about 17 marks low and gemini-3.5-flash is within
   * 2-4, so pooling them produced `real = app x 0.043 + 93.7` — a meaningless
   * line through two clusters that just says "the models differ".
   */
  const groups = {};
  for (const r of runs) {
    const key = r.board + ' — ' + (r.modelUsedForMarking || 'model not recorded');
    (groups[key] = groups[key] || []).push(r);
  }

  for (const key of Object.keys(groups)) {
    const list = groups[key];
    console.log('');
    console.log('  ' + key + ' — ' + list.length + ' anchor' + (list.length === 1 ? '' : 's'));
    for (const r of list) {
      console.log('    ' + String(r.label || r.id).slice(0, 46).padEnd(47)
        + 'app ' + String(r.totalPredicted).padStart(3) + '   real ' + String(r.totalActual).padStart(3));
    }
    const fit = fitCorrection(list.map((r) => ({ app: r.totalPredicted, real: r.totalActual })));
    console.log('');
    console.log('    fitted    : ' + describeCorrection(fit));
    console.log('    residuals : ' + fit.residuals.map((x) => 'app ' + x.app + ' -> real ' + x.real + ' (' + (x.residual >= 0 ? '+' : '') + x.residual + ')').join(', '));
    console.log('    max error : ' + fit.maxResidual + ' marks');
    if (fit.caveat) console.log('    CAVEAT    : ' + fit.caveat);
  }
  console.log('');
  console.log('  A correction is only worth applying when the anchors agree with it.');
  console.log('  Applying one fitted from two points would make the app wrong in a NEW way');
  console.log('  on every portfolio outside that range.');
  console.log('');
}

function showRun(id) {
  const store = readJson(CALIBRATION_PATH, { runs: [] });
  const record = (store.runs || []).find((r) => r.id === id);
  if (!record) fail('No calibration with id "' + id + '". Use list to see them.');
  printComparison(record);
}

/* ------------------------------------------------------------------ *
 * Entry
 * ------------------------------------------------------------------ */

async function main() {
  const args = process.argv.slice(2);
  const command = args[0];
  if (command === 'run') return runCalibration(args.slice(1));
  if (command === 'list') return listRuns();
  if (command === 'show') return showRun(args[1]);
  if (command === 'fit') return fitAll();
  if (command === 'record') return recordAnchor(args.slice(1));
  console.log('');
  console.log('  NEA Coach calibration — how close is the predicted mark to the truth?');
  console.log('');
  console.log('    node tools/calibrate.js run <pdf> --marks a=10,b=10,c=20,d=20,e=20,f=19');
  console.log('    node tools/calibrate.js record --app 95 --real 89 [--marks a=8,b=9,...]');
  console.log('    node tools/calibrate.js list');
  console.log('    node tools/calibrate.js show <id>');
  console.log('    node tools/calibrate.js fit            (fit a correction across all anchors)');
  console.log('');
  console.log('  --marks are the REAL marks awarded, out of each criterion\'s maximum.');
  console.log('  Without them a scan is just another scan, not a measurement.');
  console.log('');
}

if (require.main === module) {
  main().catch((err) => fail(String(err && err.message ? err.message : err)));
}

module.exports = { compareToKnown, compareToKnownTotal, parseMarks, buildRecord, fitCorrection, describeCorrection, checkMarksConsistent };
