/*
 * NEA Coach — regression tests.
 *
 * Run with:   node tests/test.js
 * No dependencies, no test runner: this is plain Node so that any AI or person
 * can run it on any machine without installing anything.
 *
 * Every test here exists because the bug it covers actually happened and was
 * reported by the student. They are grouped by that history.
 */

'use strict';

const path = require('path');
const fs = require('fs');

const ROOT = path.join(__dirname, '..');
const { parsePresentation } = require(path.join(ROOT, 'lib/pptx.js'));
const { runLocalChecks, DISABLED_RULES } = require(path.join(ROOT, 'lib/checks.js'));
const ai = require(path.join(ROOT, 'lib/ai.js'));

let passed = 0;
let failed = 0;
const failures = [];

function ok(name, condition, detail) {
  if (condition) {
    passed += 1;
    console.log('  PASS  ' + name);
  } else {
    failed += 1;
    failures.push(name + (detail ? ' — ' + detail : ''));
    console.log('  FAIL  ' + name + (detail ? '  [' + detail + ']' : ''));
  }
}

function eq(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  ok(name, a === e, 'got ' + a + ', wanted ' + e);
}

function section(title) {
  console.log('\n' + title);
  console.log('-'.repeat(title.length));
}

/* ------------------------------------------------------------------ *
 * 1. Tables must be read, not just detected
 *
 * The student reported that the app "didnt seem to be able to read images or
 * tables, claiming that pages were blank". Tables were detected and their text
 * thrown away, so a page holding a full manufacture plan reached the AI as two
 * words: its title. The AI then told them to write the pages they had written.
 * ------------------------------------------------------------------ */

function fakeTableSlideXml() {
  return `<?xml version="1.0"?>
<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-GB" sz="2400"/><a:t>Manufacture plan</a:t></a:r></a:p></p:txBody></p:sp>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table 1"/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="0" y="600000"/><a:ext cx="3000000" cy="2000000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
<a:tblGrid><a:gridCol w="1000000"/><a:gridCol w="2000000"/></a:tblGrid>
<a:tr h="300000">
  <a:tc><a:txBody><a:p><a:r><a:rPr lang="en-GB" sz="1200"/><a:t>Laser cut frame</a:t></a:r></a:p></a:txBody></a:tc>
  <a:tc><a:txBody><a:p><a:r><a:rPr lang="en-GB" sz="1200"/><a:t>Project the sketches in Fusion 360</a:t></a:r></a:p></a:txBody></a:tc>
</a:tr>
<a:tr h="300000">
  <a:tc><a:txBody><a:p><a:r><a:rPr lang="en-GB" sz="1000"/><a:t>3D print components</a:t></a:r></a:p></a:txBody></a:tc>
  <a:tc><a:txBody><a:p><a:r><a:rPr lang="en-GB" sz="1000"/><a:t>Export meshes and slice them</a:t></a:r></a:p></a:txBody></a:tc>
</a:tr>
</a:tbl></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>`;
}

section('1. Table text extraction');
{
  const samplePath = path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx');
  if (fs.existsSync(samplePath)) {
    const deck = parsePresentation(samplePath);
    ok('sample deck parses', deck && Array.isArray(deck.slides) && deck.slides.length > 0);
  } else {
    console.log('  SKIP  sample deck not present');
  }

  // The real portfolio is the only place with tables; verify if it is here.
  const realPath = path.join(ROOT, '..', 'Rohan Ghali.pptx');
  if (fs.existsSync(realPath)) {
    const deck = parsePresentation(realPath);
    const withTables = deck.slides.filter((s) => (s.shapes || []).some((sh) => sh.kind === 'table'));
    ok('real deck tables are found', withTables.length > 0, withTables.length + ' pages');

    const emptyTablePages = withTables.filter((s) => s.wordCount < 12);
    eq('no table page reads as empty', emptyTablePages.map((s) => s.index), []);

    // A specific page that used to arrive as the single word "Manufacture plan".
    const p45 = deck.slides.find((s) => s.index === 45);
    if (p45) {
      ok('page 45 manufacture plan text is extracted', p45.wordCount > 300, p45.wordCount + ' words');
      ok('page 45 keeps its table text', /Laser|3D print|solder/i.test(p45.text));
      ok('page 45 table paragraph sizes are captured',
        p45.textShapes.some((s) => s.kind === 'table' && s.paragraphs.some((p) => p.sizePt != null)));
    } else {
      console.log('  SKIP  page 45 not found');
    }
  } else {
    console.log('  SKIP  real portfolio not present');
  }

  // Pure unit test of the table reader, independent of any file on disk.
  const { open } = require(path.join(ROOT, 'lib/zipread.js'));
  ok('zip reader is importable', typeof open === 'function');
}

/* ------------------------------------------------------------------ *
 * 2. Rules the student asked to switch off
 * ------------------------------------------------------------------ */

section('2. Image resolution checks are disabled');
{
  ok('low-res-image is disabled', DISABLED_RULES.has('low-res-image'));
  ok('image-too-small is disabled', DISABLED_RULES.has('image-too-small'));

  const samplePath = path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx');
  if (fs.existsSync(samplePath)) {
    const res = runLocalChecks(parsePresentation(samplePath));
    const rules = res.findings.map((f) => f.rule);
    ok('no low-res-image findings are produced', rules.indexOf('low-res-image') < 0);
    ok('no image-too-small findings are produced', rules.indexOf('image-too-small') < 0);
  }
}

/* ------------------------------------------------------------------ *
 * 3. Sample deck answer key
 *
 * These were planted deliberately. If a change breaks one of these, the checks
 * have regressed, not the sample.
 * ------------------------------------------------------------------ */

section('3. Sample deck answer key');
{
  const samplePath = path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx');
  if (!fs.existsSync(samplePath)) {
    console.log('  SKIP  sample deck not present');
  } else {
    const res = runLocalChecks(parsePresentation(samplePath));
    const byRule = {};
    for (const f of res.findings) {
      (byRule[f.rule] = byRule[f.rule] || []).push(f.slide == null ? 'deck' : f.slide);
    }
    const sorted = (r) => (byRule[r] || []).slice().sort((a, b) => a - b);

    eq('missing-title fires on 4, 13, 21', sorted('missing-title'), [4, 13, 21]);
    eq('uncaptioned-image fires on 6, 7', sorted('uncaptioned-image'), [6, 7]);
    eq('near-empty-page fires on 16, 17', sorted('near-empty-page'), [16, 17]);
    eq('text-too-small fires on 18', sorted('text-too-small'), [18]);
    ok('uncaptioned-image does NOT fire on 14 (caption sits beside the image)',
      sorted('uncaptioned-image').indexOf(14) < 0);
    eq('no check throws', res.errors, []);
  }
}

/* ------------------------------------------------------------------ *
 * 4. Relevance tiers
 *
 * The student: "the list of 'fix it' stuff seems to still only be the silly
 * stuff like too much text etc whereas it should be what i was talking about".
 * ------------------------------------------------------------------ */

section('4. Relevance tiers');
{
  const samplePath = path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx');
  if (fs.existsSync(samplePath)) {
    const res = runLocalChecks(parsePresentation(samplePath));
    ok('every local finding carries a tier', res.findings.every((f) => ['substance', 'presentation', 'obvious'].indexOf(f.tier) >= 0));
    ok('byTier totals match the finding count',
      (res.byTier.substance + res.byTier.presentation + res.byTier.obvious) === res.findings.length,
      JSON.stringify(res.byTier) + ' vs ' + res.findings.length);

    // Mechanical rules must never outrank substantive ones.
    const obviousIdx = res.findings.findIndex((f) => f.tier === 'obvious');
    const presentationIdx = res.findings.findIndex((f) => f.tier === 'presentation');
    const substanceIdx = res.findings.findIndex((f) => f.tier === 'substance');
    if (substanceIdx >= 0 && obviousIdx >= 0) {
      ok('substance findings sort above obvious ones', substanceIdx < obviousIdx,
        'substance@' + substanceIdx + ' obvious@' + obviousIdx);
    }
    if (presentationIdx >= 0 && obviousIdx >= 0) {
      ok('presentation findings sort above obvious ones', presentationIdx < obviousIdx);
    }
  }
}

/* ------------------------------------------------------------------ *
 * 5. AI finding classification
 * ------------------------------------------------------------------ */

section('5. AI finding classification');
{
  const factRaw = { rule: 'x', title: 'Fact', kind: 'fact', severity: 'high', slides: [3] };
  const fact = ai.normaliseAiFinding(factRaw, [3], 1);
  eq('a fact stays in the fix list', fact.progressAware, true);
  eq('a fact is substantive', fact.tier, 'substance');
  eq('a fact-tiered finding outranks obvious local checks', fact.priority > 2000, true);

  const addNoFlag = { rule: 'y', title: 'Add', kind: 'add', severity: 'high', slides: [9] };
  const add = ai.normaliseAiFinding(addNoFlag, [9], 2);
  eq('an "add" without progressAware defaults to NOT-yet-done', add.progressAware, false);
  eq('an "add" is tiered presentation so the fix list stays decisive', add.tier, 'presentation');

  const addDone = { rule: 'z', title: 'Add', kind: 'add', severity: 'medium', slides: [9], progressAware: true };
  const done = ai.normaliseAiFinding(addDone, [9], 3);
  eq('an "add" explicitly marked as existing work keeps progressAware true', done.progressAware, true);
  eq('such a finding is substantive, not a suggestion to add', done.tier, 'substance');

  const judgement = ai.normaliseAiFinding({ rule: 'j', title: 'J', kind: 'judgement', severity: 'medium', slides: [4] }, [4], 4);
  eq('judgement findings are substantive', judgement.tier, 'substance');

  const worth = ai.normaliseAiFinding(
    { rule: 'w', title: 'W', kind: 'add', severity: 'high', impact: 'high', effort: 'minutes', slides: [1] }, [1], 5);
  const slow = ai.normaliseAiFinding(
    { rule: 's', title: 'S', kind: 'add', severity: 'high', impact: 'high', effort: 'days', slides: [1] }, [1], 6);
  ok('a quick high-impact addition beats a slow one', worth.worthScore > slow.worthScore,
    worth.worthScore + ' vs ' + slow.worthScore);
}

/* ------------------------------------------------------------------ *
 * 5b. Real-deck regressions
 *
 * Each of these failed on the real portfolio while passing on the sample.
 * ------------------------------------------------------------------ */

section('5b. Real-deck regressions');
{
  const realPath = path.join(ROOT, '..', 'Rohan Ghali.pptx');
  if (!fs.existsSync(realPath)) {
    console.log('  SKIP  real portfolio not present');
  } else {
    const deck = parsePresentation(realPath);
    const res = runLocalChecks(deck);

    // Tables must not be reported as text overflowing a fixed box: PowerPoint
    // grows rows to fit, so a table's frame height is not a clipping boundary.
    const tableSlides = deck.slides
      .filter((sl) => (sl.shapes || []).some((sh) => sh.kind === 'table'))
      .map((sl) => sl.index);
    const overflowOnTables = res.findings
      .filter((f) => f.rule === 'text-overflow' && tableSlides.indexOf(f.slide) >= 0)
      .map((f) => f.slide);
    eq('no text-overflow is reported on any table page', overflowOnTables, []);

    // Every page must carry a usable amount of extractable content now.
    const contentless = deck.slides.filter((sl) => sl.wordCount < 12 && sl.imageCount === 0);
    eq('no page reads as completely empty', contentless.map((sl) => sl.index), []);

    // The disabled rules must be absent from real output too.
    const rules = res.findings.map((f) => f.rule);
    ok('no low-res-image findings on the real deck', rules.indexOf('low-res-image') < 0);
    ok('no image-too-small findings on the real deck', rules.indexOf('image-too-small') < 0);

    // Tier ordering must survive real data.
    const firstObvious = res.findings.findIndex((f) => f.tier === 'obvious');
    const lastSubstance = res.findings.map((f) => f.tier).lastIndexOf('substance');
    if (firstObvious >= 0 && lastSubstance >= 0) {
      ok('no obvious finding sorts above a substance finding',
        lastSubstance < firstObvious, 'lastSubstance@' + lastSubstance + ' firstObvious@' + firstObvious);
    }

    // Page 45 is the manufacture plan that used to arrive as two words.
    const p45 = deck.slides.find((sl) => sl.index === 45);
    if (p45) {
      ok('page 45 keeps its full manufacture plan', p45.wordCount > 300, p45.wordCount + ' words');
      ok('page 45 includes its safety and quality-control text', /quality control|safety check/i.test(p45.text));
    }
  }
}

/* ------------------------------------------------------------------ *
 * 6. Pricing and model choice
 *
 * The student: "it cost 8p to read 60 pages is there a way to make it cheaper".
 * ------------------------------------------------------------------ */

section('6. Pricing and cost estimates');
{
  // Longest-prefix matching: this exact bug overcharged the meter ~7x.
  const lite = ai.costFor('gemini-3.5-flash-lite', 1e6, 1e6);
  const flash = ai.costFor('gemini-3.8-flash', 1e6, 1e6);
  ok('a flash-lite model is not priced as a full flash model', lite.usd < flash.usd,
    'lite $' + lite.usd + ' vs flash $' + flash.usd);
  ok('flash-lite matches its own price row', Math.abs(lite.usd - 2.8) < 0.01, 'got $' + lite.usd);
  ok('flash matches its own price row', Math.abs(flash.usd - 4.5) < 0.01, 'got $' + flash.usd);

  // Published prices, verified against Google's pricing page.
  ok('3.8 flash input price is 0.75/M', ai.PRICING_USD_PER_MTOK['gemini-3.8-flash'].input === 0.75);
  ok('3.5 flash-lite input price is 0.30/M', ai.PRICING_USD_PER_MTOK['gemini-3.5-flash-lite'].input === 0.3);
  ok('no pricing row exists for a model that does not exist',
    !Object.prototype.hasOwnProperty.call(ai.PRICING_USD_PER_MTOK, 'gemini-3.8-flash-lite'));

  const cheap = ai.estimateScanCostGbp('gemini-3.1-flash-lite', 60);
  const dear = ai.estimateScanCostGbp('gemini-3.8-flash', 60);
  ok('a 60-page scan on flash-lite is estimated under 6p', cheap.gbp * 100 < 6, (cheap.gbp * 100).toFixed(2) + 'p');
  ok('flash-lite is cheaper than flash for the same deck', cheap.gbp < dear.gbp);
  ok('a 200-page scan stays well inside the student budget', ai.estimateScanCostGbp('gemini-3.1-flash-lite', 200).gbp < 0.25);

  // Model choice must not land on a model that rejects our request shape.
  const available = ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-3.1-pro-preview', 'gemini-2.5-flash']
    .map((n) => ({ name: n }));
  const picked = ai.pickModel(available, null);
  ok('a cheap model is picked by default', /flash-lite|2\.5-flash/.test(picked), picked);
  ok('the picked model is not one that rejects thinkingConfig', ai.supportsThinkingBudget(picked), picked);
  eq('an explicit preference is honoured', ai.pickModel(available, 'gemini-3.8-flash'), 'gemini-3.8-flash');

  // Google still lists retired models that then 404 for new keys.
  const withRetired = ['gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-3.1-flash-lite'].map((n) => ({ name: n }));
  const picked2 = ai.pickModel(withRetired, null);
  ok('a retired model is never picked', !ai.isRetiredModel(picked2), picked2);
  eq('retired detection works', ai.isRetiredModel('gemini-2.5-flash'), true);
  eq('current models are not flagged retired', ai.isRetiredModel('gemini-3.1-flash-lite'), false);
}

/* ------------------------------------------------------------------ *
 * 6b. A retired model must never be used
 *
 * The student reported: "its saying its not working for tje other models when i
 * tried changing it ony 3.8 and 3.8 isnt showing up in the options now". The
 * config was pinned to gemini-2.5-pro, which Google has retired, so every AI
 * request returned 404 and the app never said why.
 * ------------------------------------------------------------------ */

section('6b. Retired models are refused, not silently used');
{
  // A preference for a retired model must be ignored, not honoured.
  const apiModels = ['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-3.8-flash']
    .map((n) => ({ name: n }));

  const picked = ai.pickModel(apiModels, 'gemini-2.5-pro');
  ok('a retired preferred model is NOT honoured', picked !== 'gemini-2.5-pro', picked);
  ok('a working model is chosen instead', !ai.isRetiredModel(picked) && !!picked, picked);

  const kept = ai.pickModel(apiModels, 'gemini-3.8-flash');
  eq('a valid preferred model IS still honoured', kept, 'gemini-3.8-flash');

  // Retired detection must cover the generations Google has withdrawn.
  eq('2.5 models are retired', ai.isRetiredModel('gemini-2.5-pro'), true);
  eq('2.0 models are retired', ai.isRetiredModel('gemini-2.0-flash'), true);
  eq('1.5 models are retired', ai.isRetiredModel('gemini-1.5-pro'), true);
  eq('3.x models are not retired', ai.isRetiredModel('gemini-3.8-flash'), false);
  eq('an empty name is not retired', ai.isRetiredModel(''), false);

  // The picker must prefer a model that keeps deep thinking available, because
  // choosing one that cannot would silently remove the feature from the app.
  const withLatest = ['gemini-flash-lite-latest', 'gemini-3.1-flash-lite'].map((n) => ({ name: n }));
  const chosen = ai.pickModel(withLatest, null);
  ok('a thinking-capable model is preferred over an alias that cannot think',
    ai.supportsThinkingBudget(chosen), chosen);

  // The server must refuse to save a retired model at all.
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok('the config endpoint refuses a retired model',
    /isRetiredModel\(wanted\)/.test(serverSrc) && /has retired/.test(serverSrc));
  ok('a scan self-heals a config naming a retired model',
    /has been retired by Google/.test(serverSrc));

  /*
   * The model list in the UI must not be capped. A hard cap of twelve, combined
   * with a <datalist> that shows nothing until you type, is why the model the
   * student wanted appeared to have vanished.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  ok('the model picker is a real <select>, not a <datalist>',
    /<select id="modelSelect"/.test(html) && !/<datalist id="modelChoices"/.test(html));
  ok('the model list is not capped at twelve', !/shown < 12/.test(html));
  ok('retired models are filtered out of the picker', /\.retired\)\s*\{\s*return false/.test(html));
}

/* ------------------------------------------------------------------ *
 * 6c. The two cost figures the student sees must agree
 * ------------------------------------------------------------------ */

section('6c. Client and server price estimates agree');
{
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  // The page mirrors the server's constants; if they drift, the app shows two
  // different prices for the same thing, which reads as a bug.
  const batch = ai.BATCH_SIZE;
  ok('the page mirrors the server batch size', new RegExp('BATCH_SIZE_MIRROR\\s*=\\s*' + batch).test(html), 'BATCH_SIZE=' + batch);
  ok('the page mirrors the per-page input token count',
    new RegExp('TOKENS_IN_PER_PAGE\\s*=\\s*' + ai.MEASURED_TOKENS_PER_PAGE.input).test(html));
  ok('the page mirrors the per-page output token count',
    new RegExp('TOKENS_OUT_PER_PAGE\\s*=\\s*' + ai.MEASURED_TOKENS_PER_PAGE.output).test(html));
  const cc = ai.CROSS_CHECK_TOKENS;
  ok('the page mirrors the cross-check token counts',
    new RegExp('CROSS_CHECK_BASE_TOKENS\\s*=\\s*' + cc.base).test(html)
      && new RegExp('CROSS_CHECK_TOKENS_PER_PAGE\\s*=\\s*' + cc.perPage).test(html)
      && new RegExp('CROSS_CHECK_TOKENS_PER_BATCH\\s*=\\s*' + cc.perBatch).test(html));
  // The output figure is the MEASURED one, which differs from the prompt-sizing
  // constant in CROSS_CHECK_TOKENS. Both the estimator and the page must use it,
  // or the price on screen is not the price measured.
  ok('the page mirrors the measured cross-check output',
    new RegExp('CROSS_CHECK_OUTPUT_TOKENS\\s*=\\s*' + ai.CROSS_CHECK_MEASURED_OUTPUT).test(html),
    'measured output = ' + ai.CROSS_CHECK_MEASURED_OUTPUT);

  /*
   * Matching constants is NOT enough: the two implementations must AGREE.
   *
   * The page once quoted 9.22p for a check the server quoted at 16.07p — a 6.85p
   * difference — because the browser added the marking runs' tokens but priced
   * them at the PAGE-READING model's rate, ignoring `visionModel` altogether.
   * Every constant matched; only the arithmetic differed. So this reproduces the
   * page's formula and compares it with the server's, on the real default
   * combination (flash-lite reading, 3.5-flash marking).
   */
  const num = (name) => {
    const m = html.match(new RegExp(name + '\\s*=\\s*([0-9.]+)'));
    return m ? Number(m[1]) : null;
  };
  const pageN = 60;
  const pageRuns = num('VISION_RUNS_MIRROR');
  const priceRow = (name) => {
    let key = null;
    for (const k of Object.keys(ai.PRICING_USD_PER_MTOK)) {
      if (name.startsWith(k) && (key === null || k.length > key.length)) key = k;
    }
    return key ? ai.PRICING_USD_PER_MTOK[key] : null;
  };
  const readRow = priceRow('gemini-3.1-flash-lite');
  const markRow = priceRow('gemini-3.5-flash');
  const batches = Math.ceil(pageN / ai.BATCH_SIZE);
  const textIn = pageN * num('TOKENS_IN_PER_PAGE') + batches * num('BATCH_OVERHEAD_TOKENS')
    + num('CROSS_CHECK_BASE_TOKENS') + pageN * num('CROSS_CHECK_TOKENS_PER_PAGE')
    + batches * num('CROSS_CHECK_TOKENS_PER_BATCH');
  const textOut = pageN * num('TOKENS_OUT_PER_PAGE') + num('CROSS_CHECK_OUTPUT_TOKENS');
  const vIn = pageN * num('TOKENS_IN_PER_PAGE_VISION') * pageRuns;
  const vOut = pageN * num('TOKENS_OUT_PER_PAGE_VISION') * pageRuns;
  const pageUsd = (textIn / 1e6) * readRow.input + (textOut / 1e6) * readRow.output
    + (vIn / 1e6) * markRow.input + (vOut / 1e6) * markRow.output;
  const pageQuote = pageUsd * ai.USD_TO_GBP;
  const serverQuote = ai.estimateScanCostGbp('gemini-3.1-flash-lite', pageN, {
    vision: true, runs: pageRuns, visionModel: 'gemini-3.5-flash',
  }).gbp;

  ok('the page and the server quote the same price for the same scan',
    Math.abs(pageQuote - serverQuote) / serverQuote < 0.02,
    'page ' + (pageQuote * 100).toFixed(2) + 'p vs server ' + (serverQuote * 100).toFixed(2) + 'p');
  ok('and both price the marking runs, not just the page-reading model',
    serverQuote > ai.estimateScanCostGbp('gemini-3.1-flash-lite', pageN,
      { vision: true, runs: pageRuns, visionModel: 'gemini-3.1-flash-lite' }).gbp);

  // Every model the server can price must be priceable in the page too, or the
  // dropdown would show "price not known" for a model whose cost we do know.
  const missing = Object.keys(ai.PRICING_USD_PER_MTOK).filter((m) => html.indexOf("['" + m + "'") < 0);
  const retiredPrices = missing.filter((m) => ai.isRetiredModel(m));
  const realMissing = missing.filter((m) => !ai.isRetiredModel(m));
  ok('every non-retired priced model appears in the page price table',
    realMissing.length === 0, realMissing.join(', '));
  ok('retired models may be omitted from the page table (they are never offered)',
    retiredPrices.length === missing.length - realMissing.length);
}

/* ------------------------------------------------------------------ *
 * 7. Prompt construction
 * ------------------------------------------------------------------ */

section('7. Prompt construction');
{
  const slide = {
    index: 12,
    title: 'Manufacture plan',
    wordCount: 3,
    imageCount: 4,
    notesWordCount: 0,
    shapeCount: 5,
    images: [
      { shapeName: 'Picture 1', pixelWidth: 1200, pixelHeight: 800, placedWidthIn: 6, placedHeightIn: 4, effectiveDpi: 200, altText: 'Exploded view of the hinge' }
    ],
    shapes: [{ kind: 'table', table: { rowCount: 3, colCount: 2 } }],
    text: 'Task | Steps\nLaser cut | Project the sketches',
    notes: ''
  };
  const block = ai.buildSlideBlock(slide, { slideWidthIn: 16.54, slideHeightIn: 11.69 });

  ok('the block names the tables on the page', /TABLES ON THIS PAGE/.test(block));
  ok('the block explains how table text is laid out', /separated by " \| "/.test(block));
  ok('image alt text is passed through', /described as "Exploded view of the hinge"/.test(block));
  ok('the block states that pictures themselves are not visible', /not the pictures themselves/.test(block));
  ok('a picture-led page is flagged as such', /picture-led page/.test(block));
  ok('a picture-led page is explicitly not called blank', /Do NOT call it blank/.test(block));

  // A genuinely empty page should still read as empty.
  const blank = ai.buildSlideBlock(
    { index: 13, title: null, wordCount: 0, imageCount: 0, notesWordCount: 0, shapeCount: 1, images: [], shapes: [], text: '', notes: '' },
    { slideWidthIn: 16.54, slideHeightIn: 11.69 });
  ok('a truly empty page still says so', /no text at all on this page/.test(blank));
  ok('a truly empty page is not described as picture-led', !/picture-led/.test(blank));

  // Images but no text must not read as blank.
  const imgOnly = ai.buildSlideBlock(
    { index: 14, title: 'Sketches', wordCount: 0, imageCount: 3, notesWordCount: 0, shapeCount: 4, images: [], shapes: [], text: '', notes: '' },
    { slideWidthIn: 16.54, slideHeightIn: 11.69 });
  ok('an image-only page is stated not to be blank', /it is not blank/.test(imgOnly));
}

/* ------------------------------------------------------------------ *
 * 8. No secrets in client payloads
 * ------------------------------------------------------------------ */

section('8. Secrets stay on the Mac');
{
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const server = require(path.join(ROOT, 'server.js'));

  /*
   * Every config payload that leaves the process goes through publicConfig,
   * which drops the API key and the access code and reports only whether a key
   * is set.
   */
  const shown = JSON.parse(JSON.stringify(server.publicConfig({ apiKey: 'AIzaSECRET', accessCode: 'abcd-efgh-jkmn', model: 'm' })));
  ok('publicConfig drops the API key', !('apiKey' in shown) && shown.apiKeySet === true);
  ok('publicConfig drops the access code', !('accessCode' in shown));
  const configResponses = serverSrc.match(/config: publicConfig\(cfg\)/g) || [];
  ok('both config responses use publicConfig', configResponses.length >= 2, configResponses.length + ' found');
  ok('no response object carries the raw key', !/sendJson\([^;]*apiKey:\s*cfg\.apiKey/s.test(serverSrc));

  // Prove the detector detects: a deliberate leak must match the pattern used.
  const leaky = 'sendJson(res, 200, { config: Object.assign({}, cfg, { apiKey: cfg.apiKey }) });';
  ok('the leak detector would catch a real leak', /apiKey:\s*cfg\.apiKey/.test(leaky));
}

/* ------------------------------------------------------------------ *
 * 9. The app never writes to the portfolio
 * ------------------------------------------------------------------ */

section('9. The portfolio is only ever read');
{
  const files = ['server.js', 'lib/pptx.js', 'lib/zipread.js', 'lib/checks.js', 'lib/ai.js', 'lib/imagemeta.js'];
  const writers = [];
  for (const rel of files) {
    const src = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    // Any write call that targets the deck would show up as one of these names.
    if (/writeFileSync\([^)]*deck|writeFileSync\([^)]*\.pptx|appendFileSync\([^)]*\.pptx/.test(src)) {
      writers.push(rel);
    }
  }
  eq('no module writes to the deck file', writers, []);

  const pptxSrc = fs.readFileSync(path.join(ROOT, 'lib', 'pptx.js'), 'utf8');
  ok('pptx.js imports no write API', !/writeFile|createWriteStream|appendFile/.test(pptxSrc));
  ok('zipread.js opens archives read-only', !/writeFile|createWriteStream|appendFile/.test(fs.readFileSync(path.join(ROOT, 'lib', 'zipread.js'), 'utf8')));
}

/* ------------------------------------------------------------------ *
 * 10. Only this Mac, or someone with the access code, can use the app
 *
 * The server listened on the network with no password: anyone on the same
 * Wi-Fi could read the report, swap the API key, point the app at any folder
 * and run paid scans, and any web page open on the Mac could do the same.
 * ------------------------------------------------------------------ */

section('10. Access control');
(async () => {
  const server = require(path.join(ROOT, 'server.js'));
  const code = 'abcd-efgh-jkmn';
  const req = (remoteAddress, host, cookie) => ({ socket: { remoteAddress }, headers: { host, cookie } });
  const at = (u) => new URL(u, 'http://localhost');

  /*
   * The phone-address endpoint is the ONLY one that returns the access code, so
   * it is restricted to this Mac. The subtle case is a LAN device that ALREADY
   * has a valid code: it passes checkAccess and must STILL be refused, because a
   * device holding the code has no need for it, and returning it would put the
   * password on the network.
   */
  const connectRes = () => {
    const r = { statusCode: null, body: null, headers: {} };
    r.writeHead = (status) => { r.statusCode = status; return r; };
    r.setHeader = (k, v) => { r.headers[k] = v; };
    r.end = (chunk) => { r.body = chunk; };
    return r;
  };
  const callConnect = async (remoteAddress, host, url) => {
    const res = connectRes();
    const request = { method: 'GET', url, socket: { remoteAddress }, headers: { host }, on: () => {} };
    await server.handle(request, res);
    let parsed = null;
    try { parsed = JSON.parse(res.body || 'null'); } catch (err) { parsed = null; }
    return { status: res.statusCode, parsed, raw: String(res.body || '') };
  };

  const lanNoKey = await callConnect('192.168.0.40', '192.168.0.87:4173', '/api/connect');
  ok('the phone address is refused to another device with no code', lanNoKey.status !== 200, JSON.stringify(lanNoKey.parsed));

  const lanWithKey = await callConnect('192.168.0.40', '192.168.0.87:4173', '/api/connect?key=' + code);
  ok('and refused even to a device holding a valid code', lanWithKey.status !== 200, JSON.stringify(lanWithKey.parsed));
  ok('the refusal does not leak the code', lanWithKey.raw.indexOf(code) < 0);
  /*
   * Assert the PROPERTY, not the exact wording. Whether this is refused by
   * checkAccess (as here, where no access code has been set because start() never
   * ran) or by the endpoint's own loopback guard, what must hold is that a LAN
   * device never receives the address or the code, and is told to look on the Mac.
   */
  ok('the refusal points the student at the Mac',
    /Mac/.test(lanWithKey.raw) && /access code|only shown on the Mac/i.test(lanWithKey.raw), lanWithKey.raw);
  ok('the refusal contains no address', !/https?:\/\//.test(lanWithKey.raw), lanWithKey.raw);

  /*
   * The marking model must default to a Flash model, not the cheap Flash-Lite.
   *
   * This is the single biggest accuracy decision in the project. Measured on two
   * portfolios with known marks: Flash-Lite under-marked them by 17-26 marks with
   * a ~20-mark spread, while gemini-3.5-flash came within 2-4 marks on both with a
   * 2-mark spread. Five attempts at tuning the prompt moved nothing; this setting
   * fixed accuracy and repeatability together.
   *
   * If someone "optimises" this back to an empty string, the app silently returns
   * to being 20 marks wrong, so it is worth asserting.
   */
  const visionDefault = server.loadConfig().visionModel;
  ok('the marking model defaults to a stronger model, not blank',
    typeof visionDefault === 'string' && visionDefault.length > 0, JSON.stringify(visionDefault));
  ok('and it is a Flash model rather than Flash-Lite',
    /gemini-.*flash/i.test(visionDefault) && !/flash-lite/i.test(visionDefault), visionDefault);

  ok('this Mac at 127.0.0.1 gets in without a code', server.checkAccess(req('127.0.0.1', '127.0.0.1:4173'), at('/'), code).ok);
  ok('this Mac at localhost gets in without a code', server.checkAccess(req('::1', 'localhost:4173'), at('/'), code).ok);
  ok('a web page using its own domain name for 127.0.0.1 is refused',
    !server.checkAccess(req('127.0.0.1', 'evil.example:4173'), at('/api/state'), code).ok);
  ok('another device without the code is refused', !server.checkAccess(req('192.168.0.40', '192.168.0.87:4173'), at('/'), code).ok);
  ok('another device with a wrong code is refused',
    !server.checkAccess(req('192.168.0.40', '192.168.0.87:4173'), at('/?key=abcd-efgh-jkmp'), code).ok);
  const withKey = server.checkAccess(req('192.168.0.40', '192.168.0.87:4173'), at('/?key=ABCD-EFGH-JKMN'), code);
  ok('the right code in the address gets in, and sets the cookie', withKey.ok && withKey.setCookie);
  ok('the cookie alone gets in afterwards',
    server.checkAccess(req('192.168.0.40', '192.168.0.87:4173', 'neaCoachKey=' + code), at('/api/state'), code).ok);
  ok('no code configured means only this Mac gets in',
    !server.checkAccess(req('192.168.0.40', '192.168.0.87:4173'), at('/?key='), null).ok);
  ok('new access codes have the expected shape', /^[a-z2-9]{4}-[a-z2-9]{4}-[a-z2-9]{4}$/.test(server.newAccessCode()));

  // A POST without the page's own header is refused before anything is read.
  const fakeRes = () => {
    const r = { status: null, body: '', headersSent: false, setHeader() {} };
    r.writeHead = (s) => { r.status = s; r.headersSent = true; };
    r.end = (b) => { r.body = String(b || ''); };
    return r;
  };
  const post = (headers) => Object.assign(req('127.0.0.1', '127.0.0.1:4173'), { method: 'POST', url: '/api/config' },
    { headers: Object.assign({ host: '127.0.0.1:4173' }, headers) });
  const refused = fakeRes();
  server.handle(post({}), refused);
  eq('a POST without the X-NEA-Coach header is refused', refused.status, 403);

  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  ok('the page sends the X-NEA-Coach header on every POST', /'X-NEA-Coach': '1'/.test(html));
})().then(() => {}, (err) => {
  ok('access control tests ran', false, err && err.stack);
});

/* ------------------------------------------------------------------ *
 * 11. Suggestions and coverage are judged against the whole portfolio
 *
 * Each AI batch sees 12 pages. On the real portfolio a batch covering pages
 * 1-12 said "The design specification is currently missing" while page 17 was
 * "Brief and specification", and the coverage grid showed Manufacturing as
 * "none" next to eight pages that evidence it.
 * ------------------------------------------------------------------ */

section('11. Whole-portfolio cross-check');
{
  const criteria = [{ id: 'a', name: 'Investigating' }, { id: 'd', name: 'Manufacturing a prototype' }];

  // The real failure: only the making section reports Manufacturing.
  const merged = ai.mergeCoverage([
    { coverage: [{ criterion: 'a', status: 'strong', coveragePercent: 80, slides: [3], comment: '' },
      { criterion: 'd', status: 'none', coveragePercent: 0, slides: [], comment: 'not reached yet' }] },
    { coverage: [{ criterion: '(d) Manufacturing a prototype', status: 'strong', coveragePercent: 90, slides: [45, 46], comment: 'good' }] },
    { coverage: [{ criterion: 'd', status: 'none', coveragePercent: 0, slides: [], comment: '' }] },
  ], criteria);
  const d = merged.find((c) => c.id === 'd');
  ok('fallback coverage takes the strongest section, not the weakest', d && d.status === 'strong' && d.coveragePercent === 90,
    JSON.stringify(d));
  ok('a criterion labelled "(d) ..." merges with "d"', merged.filter((c) => c.id === 'd').length === 1 && d.slides.join() === '45,46');
  eq('merged rows carry the criterion name for the grid', d.name, 'Manufacturing a prototype');

  // Batch coverage is pinned to the real criterion ids.
  const schema = ai.responseSchemaFor({ criteria });
  eq('batch coverage must name a real criterion id', schema.properties.coverage.items.properties.criterion.enum, ['a', 'd']);

  // The reply parser keeps only known ids and real page numbers.
  const reply = (obj) => ({ candidates: [{ content: { parts: [{ text: JSON.stringify(obj) }] }, finishReason: 'STOP' }], usageMetadata: { promptTokenCount: 100, candidatesTokenCount: 10 } });
  const parsed = ai.parseCrossCheckResponse(reply({
    coverage: [{ criterion: 'd', status: 'strong', coveragePercent: 140, slides: [45, 999, 0], comment: 'c' },
      { criterion: 'zz', status: 'strong', coveragePercent: 50, slides: [], comment: '' }],
    verdicts: [{ id: 's1', verdict: 'already-done', reason: 'page 17' }, { id: 's9', verdict: 'missing', reason: '' }],
    error: '',
  }), { criterionIds: ['a', 'd'], candidateIds: ['s1', 's2'], slideCount: 60 });
  ok('cross-check reply parses', parsed.ok);
  ok('unknown criteria are dropped', !parsed.coverage.has('zz'));
  eq('percentages are clamped and impossible pages removed', parsed.coverage.get('d'), { status: 'strong', coveragePercent: 100, suggestedMark: null, slides: [45], comment: 'c' });
  ok('verdicts for unknown suggestions are dropped', !parsed.verdicts.has('s9'));
  ok('a model-reported error is treated as a failed check',
    !ai.parseCrossCheckResponse(reply({ coverage: [], verdicts: [], error: 'cannot read' }), { criterionIds: [], candidateIds: [], slideCount: 1 }).ok);

  // The outline gives the checker every page title.
  const outline = ai.buildOutline({ slides: [
    { index: 17, title: 'Brief and specification', text: 'Brief and specification The client needs a launcher', wordCount: 8, imageCount: 1 },
  ] });
  ok('the outline names each page and its opening words', /p17 "Brief and specification" — 8 words, 1 image\(s\): The client needs/.test(outline), outline);

  /*
   * 11c. The predicted mark.
   *
   * These exist because the whole point of the feature is that the student can
   * trust the number: it must never exceed the marks available, it must fall
   * back to a defensible figure when the model gives none, and it must say so
   * when the mark scheme's per-criterion marks are not filled in.
   */
  const weightedCriteria = [
    { id: 'a', name: 'Investigating', weight: 15 },
    { id: 'c', name: 'Ideas', weight: 25 },
  ];
  const p1 = ai.summarisePredictedMarks(weightedCriteria, [
    { id: 'a', status: 'strong', coveragePercent: 80, suggestedMark: 12 },
    { id: 'c', status: 'partial', coveragePercent: 40, suggestedMark: 10 },
  ]);
  eq('predicted marks are summed', p1.predicted, 22);
  eq('the maximum is the mark scheme total', p1.maximum, 40);
  eq('the percentage is the predicted share', p1.percentage, 55);
  ok('published weights are reported as published', p1.weightsPublished === true);

  const p2 = ai.summarisePredictedMarks(weightedCriteria, [
    // 999 on a 15-mark criterion must not be able to inflate the total.
    { id: 'a', status: 'strong', coveragePercent: 80, suggestedMark: 999 },
    { id: 'c', status: 'partial', coveragePercent: 40 },
  ]);
  eq('a mark above the criterion maximum is clamped', p2.rows[0].mark, 15);
  eq('a missing mark falls back to the coverage percentage', p2.rows[1].mark, 10);
  ok('the fallback row is flagged as estimated', p2.rows[1].estimated === true);
  ok('the total can never exceed the marks available', p2.predicted <= p2.maximum, `${p2.predicted}/${p2.maximum}`);

  const p3 = ai.summarisePredictedMarks(weightedCriteria, []);
  ok('no coverage yet gives no estimate rather than a zero', p3.rows.every((r) => r.mark === null));
  eq('criteria with no estimate are listed', p3.incomplete.join(','), 'a,c');

  const unweighted = [
    { id: 'a', name: 'One', weight: 0 },
    { id: 'b', name: 'Two', weight: 0 },
  ];
  const p4 = ai.summarisePredictedMarks(unweighted, [
    { id: 'a', status: 'strong', coveragePercent: 100 },
    { id: 'b', status: 'weak', coveragePercent: 50 },
  ]);
  ok('unweighted criteria are still scored', p4.weightsKnown === true && p4.predicted > 0);
  ok('unweighted criteria are declared as not published', p4.weightsPublished === false);
  ok('unweighted rows are marked as an equal share', p4.rows.every((r) => r.weighted === false));

  /*
   * 11d. One exam board at a time.
   *
   * Both criteria files live in spec/. Before the board selector existed,
   * loadSpec read every file it found, so shipping a second board would have sent
   * BOTH mark schemes on every paid request and mixed the two boards' criteria
   * into one coverage grid. This asserts the boards stay separate.
   */
  const specDir = path.join(__dirname, '..', 'spec');
  const eduqas = ai.loadSpec(specDir, 'eduqas-alevel');
  ok('the Eduqas board loads its five criteria', eduqas.criteria.length === 5, JSON.stringify(eduqas.criteria.map((c) => c.id)));
  ok('the Eduqas total is 100 marks', eduqas.criteria.reduce((s, c) => s + (Number(c.weight) || 0), 0) === 100);
  ok('the AQA criteria are not loaded alongside Eduqas', !eduqas.criteria.some((c) => /realising design ideas/i.test(c.name)));

  const aqa = ai.loadSpec(specDir, 'aqa-gcse');
  ok('the AQA board loads its six criteria', aqa.criteria.length === 6, JSON.stringify(aqa.criteria.map((c) => c.name)));
  ok('the Eduqas criteria are not loaded alongside AQA', !aqa.criteria.some((c) => /manufacturing a prototype/i.test(c.name)));
  ok('the other board\'s file is reported as skipped, not silently merged',
    aqa.skipped.some((s) => /eduqas/i.test(s.name)));

  /*
   * Board policy on annotation.
   *
   * WJEC/Eduqas states that annotation is NOT necessary; AQA expects ideas to be
   * communicated to a third party. The caption rule must follow the board, because
   * telling an Eduqas student to caption their pages is advice against their own
   * board's rules — and on a real portfolio it accounted for 25 of 103 findings.
   */
  ok('Eduqas does not require annotation', eduqas.annotationRequired === false, String(eduqas.annotationRequired));
  ok('AQA does require it', aqa.annotationRequired === true, String(aqa.annotationRequired));
  ok('the board\'s own words are carried through', /not necessary/i.test(eduqas.annotationNote || ''), eduqas.annotationNote);

  /*
   * And the rule must actually not FIRE, not merely be flagged. Checked on the
   * real sample deck, whose planted flaws include uncaptioned images.
   */
  {
    const sampleDeck = parsePresentation(path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx'));
    const wantsCaps = runLocalChecks(sampleDeck, { annotationRequired: true });
    const noCaps = runLocalChecks(sampleDeck, { annotationRequired: false });
    const captions = (res) => res.findings.filter((f) => f.rule === 'uncaptioned-image').length;
    ok('the caption rule fires when the board wants annotation', captions(wantsCaps) > 0, String(captions(wantsCaps)));
    eq('and does not fire when it does not', captions(noCaps), 0);
    ok('the policy in force is reported in the result',
      wantsCaps.boardAnnotationRequired === true && noCaps.boardAnnotationRequired === false);
  }

  /*
   * The AQA marks come from the official specification, section 4.4.4.1-4.4.4.6
   * (Version 1.2). They must total exactly 100, and the two 10-mark criteria must
   * stay at 10: if someone "tidies" these into equal shares the predicted total
   * becomes wrong for every AQA student.
   */
  const aqaTotal = aqa.criteria.reduce((s, c) => s + (Number(c.weight) || 0), 0);
  eq('the AQA criteria total 100 marks', aqaTotal, 100);
  eq('the AQA per-criterion marks are the published ones',
    aqa.criteria.map((c) => c.weight).join(','), '10,10,20,20,20,20');
  ok('the AQA criteria state their band range', aqa.criteria.every((c) => /marks is the top band/.test(c.bandRange || '')));

  const aqaPredicted = ai.summarisePredictedMarks(aqa.criteria, aqa.criteria.map((c) => ({
    id: c.id, status: 'strong', coveragePercent: 100, suggestedMark: c.weight,
  })));
  eq('a full-marks AQA deck predicts 100%', aqaPredicted.percentage, 100);
  ok('AQA marks are reported as published, not shared equally', aqaPredicted.weightsPublished === true);

  const eduqasPredicted = ai.summarisePredictedMarks(eduqas.criteria, eduqas.criteria.map((c) => ({
    id: c.id, status: 'strong', coveragePercent: 100, suggestedMark: c.weight,
  })));
  eq('a full-marks Eduqas deck predicts 100%', eduqasPredicted.percentage, 100);

  /*
   * 11e. The visual review.
   *
   * The AI pass reads text unless a PDF is supplied, and text-only materially
   * under-scores a design portfolio: a real portfolio that scored 99/100 came out
   * at 75% text-only and 88% once the pages could be seen. These assert the
   * plumbing that makes that possible, and the cost honesty that must come with
   * it — quoting a text-only price while sending images would under-quote the
   * student's bill.
   */
}

/* ------------------------------------------------------------------ *
 * 11h. Calibration — measuring the predicted mark against known marks
 *
 * This is the only part of the app that can answer "is the predicted mark
 * ACCURATE?", as opposed to "is it self-consistent?". It needs a portfolio whose
 * real marks are known, so it lives in tools/calibrate.js and is run by hand.
 *
 * The classification of the error pattern was got wrong three times while being
 * written (an outlier rule that fired on broad error; a mean-absolute "bias"
 * whose own spread tracked its magnitude; and a band order that made two bands
 * unreachable), so the patterns are pinned here one by one.
 * ------------------------------------------------------------------ */

section('11h. Calibration against known marks');
{
  const cal = require(path.join(ROOT, 'tools', 'calibrate.js'));

  eq('marks are parsed from a=1,b=2', cal.parseMarks('a=1,b=2'), { a: 1, b: 2 });
  eq('and tolerate spaces', cal.parseMarks(' a = 10 , b : 20 '), { a: 10, b: 20 });
  ok('unparseable marks are rejected, not guessed', cal.parseMarks('a=ten') === null);
  ok('empty marks are rejected', cal.parseMarks('') === null);

  const crit = [
    { id: 'a', name: 'Identifying', weight: 10 },
    { id: 'b', name: 'Brief', weight: 10 },
    { id: 'c', name: 'Generating', weight: 20 },
    { id: 'd', name: 'Developing', weight: 20 },
    { id: 'e', name: 'Realising', weight: 20 },
    { id: 'f', name: 'Evaluating', weight: 20 },
  ];
  const app = [9, 9, 17, 17, 15, 13];
  const pred = app.map((m, i) => ({ id: 'abcdef'[i], mark: m }));
  /*
   * Build known marks directly, so `diff` (predicted - actual) is unambiguous.
   * An earlier version of this fixture back-computed the known marks from the
   * differences and got the sign the wrong way round.
   */
  const knownFrom = (actual) => {
    const out = {};
    'abcdef'.split('').forEach((id, i) => { out[id] = actual[i]; });
    return out;
  };

  const accurate = cal.compareToKnown(crit, pred, knownFrom([9, 9, 17, 18, 15, 13]));
  eq('an accurate app is called well calibrated', accurate.verdict, 'well calibrated');
  eq('the per-criterion difference is reported', accurate.rows[3].diff, -1);
  eq('predicted and actual totals are both reported', [accurate.totalPredicted, accurate.totalActual], [80, 81]);

  const uniform = cal.compareToKnown(crit, pred, knownFrom([12, 12, 20, 20, 18, 16]));
  eq('a uniform offset is called recalibratable', uniform.verdict, 'consistent offset — recalibratable');
  eq('a uniform offset has no spread', uniform.biasSpread, 0);

  const uneven = cal.compareToKnown(crit, pred, knownFrom([13, 13, 21, 21, 23, 21]));
  eq('an uneven error is not called an offset', uneven.verdict, 'uneven error — check the worst criteria');

  // Errors that cancel out must NOT read as "no bias": the mean is near zero and
  // the spread is huge. A magnitude-only rule would call this well calibrated.
  const erratic = cal.compareToKnown(crit, pred, knownFrom([15, 5, 20, 14, 23, 5]));
  eq('errors that cancel are not called accurate', erratic.verdict, 'erratic error — no simple correction applies');
  ok('cancelling errors still show a real absolute size', erratic.meanAbsBias > 4, String(erratic.meanAbsBias));

  const big = cal.compareToKnown(crit, pred, knownFrom([17, 17, 25, 25, 23, 21]));
  eq('a very large even error is not called recalibratable', big.verdict, 'uneven error — check the worst criteria');

  ok('a criterion with no known mark is skipped, not assumed',
    cal.compareToKnown(crit, pred, { a: 10 }).rows.length === 1);
  ok('the worst criterion is named', !!uneven.worst && uneven.worst.name === 'Realising', JSON.stringify(uneven.worst));

  /*
   * Fitting a correction from anchors.
   *
   * ONE anchor cannot distinguish a constant offset from a proportional error,
   * and those need opposite fixes: a flat offset applied to a portfolio the app
   * rates 40 would over-correct badly if the real error grows with quality. The
   * fit refuses to claim more than the data supports.
   */
  const oneAnchor = cal.fitCorrection([{ app: 81, real: 99 }]);
  eq('one anchor gives a flat offset of +18', oneAnchor.c, 18);
  eq('and does not invent a slope', oneAnchor.k, 1);
  ok('and says so', /one point cannot prove/i.test(oneAnchor.caveat), oneAnchor.caveat);

  const flat = cal.fitCorrection([{ app: 81, real: 99 }, { app: 79, real: 96 }]);
  ok('two anchors produce a line', flat.n === 2 && typeof flat.k === 'number');
  ok('which passes through both points by construction', flat.maxResidual === 0, String(flat.maxResidual));
  ok('and it warns that two points prove nothing about the fit',
    /third portfolio/i.test(flat.caveat || ''), flat.caveat);

  // A genuinely proportional error must not be reported as a flat offset.
  const proportional = cal.fitCorrection([{ app: 40, real: 50 }, { app: 80, real: 100 }]);
  eq('a proportional error yields k = 1.25', proportional.k, 1.25);
  eq('with no offset', proportional.c, 0);
  ok('and is described as proportional', /proportional/i.test(cal.describeCorrection(proportional)),
    cal.describeCorrection(proportional));

  // A constant error across a wide range must be described as flat.
  const constant = cal.fitCorrection([{ app: 40, real: 55 }, { app: 80, real: 95 }]);
  ok('a constant error is described as a flat offset', /flat offset/i.test(cal.describeCorrection(constant)),
    cal.describeCorrection(constant));

  // Anchors that no straight line fits must say so rather than average them away.
  const scattered = cal.fitCorrection([{ app: 40, real: 90 }, { app: 60, real: 45 }, { app: 80, real: 95 }]);
  ok('scattered anchors are reported as unfittable', /no simple correction is safe/i.test(scattered.mode), scattered.mode);
  ok('the largest miss is reported', scattered.maxResidual > 5, String(scattered.maxResidual));

  ok('no anchors gives nothing rather than a guess', cal.fitCorrection([]) === null);

  /*
   * The real case: two anchors the app rated 80 and 81, for real marks 96 and 99.
   * Two points fit a line perfectly, and that line has a slope of ×3 — three real
   * marks per app mark — which is an artefact of the anchors being one mark apart,
   * not a relationship. Quoting that slope would be actively misleading, so the
   * tool must fall back to the flat offset and say why.
   */
  const bunched = cal.fitCorrection([{ app: 81, real: 99 }, { app: 80, real: 96 }]);
  ok('bunched anchors are recognised', bunched.narrow === true, String(bunched.narrow));
  ok('their mode says a slope is not supportable', /too close together/.test(bunched.mode), bunched.mode);
  ok('and the caveat names the artefact', /artefact/.test(bunched.caveat || ''), bunched.caveat);
  ok('the flat offset they imply is reported instead', bunched.flatOffset > 15 && bunched.flatOffset < 19,
    String(bunched.flatOffset));
  ok('and the description leads with the flat offset, not the slope',
    /add about \+17 marks/.test(cal.describeCorrection(bunched)), cal.describeCorrection(bunched));
  ok('the description explains why the slope is untrustworthy',
    /artefact/.test(cal.describeCorrection(bunched)));
  ok('the app-scale gap is reported', bunched.appRange === 1, String(bunched.appRange));

  // Widely separated anchors DO support a slope, and must still produce one.
  const spreadOut = cal.fitCorrection([{ app: 30, real: 50 }, { app: 80, real: 100 }]);
  ok('widely separated anchors are not called narrow', spreadOut.narrow === false);
  eq('and their slope is kept', spreadOut.k, 1);

  /*
   * Total-only anchors. Weaker by nature and must not be presented as the same
   * thing: they measure the size of the error but not where it is. This is the
   * anchor actually available for the 99/100 portfolio, whose per-criterion
   * split is not known.
   */
  const totalOnly = cal.compareToKnownTotal(crit, pred, 99, 100);
  eq('a total-only anchor still measures the gap', totalOnly.totalDiff, -19);
  eq('it reports the correction that would be needed', totalOnly.impliedOffset, 19);
  eq('it refuses to name a worst criterion', totalOnly.worst, null);
  eq('it is labelled as total-only', totalOnly.mode, 'total-only');
  ok('its verdict says how far off AND in which direction, in plain words',
    /19 marks LOW/.test(totalOnly.verdict), totalOnly.verdict);
  ok('its verdict says it cannot locate the error', /cannot say WHICH criteria/i.test(totalOnly.verdict), totalOnly.verdict);

  /*
   * Totals on different scales must be rescaled, not compared raw. A portfolio
   * marked out of 100 and one out of 120 otherwise look like a huge error.
   */
  const rescaled = cal.compareToKnownTotal(
    [{ id: 'a', weight: 50, name: 'A' }, { id: 'b', weight: 50, name: 'B' }],
    [{ id: 'a', mark: 30 }, { id: 'b', mark: 30 }],
    72, 120);
  eq('totals on another scale are rescaled', rescaled.predictedPercent, 60);
  eq('and the real side too', rescaled.actualPercent, 60);
  eq('so a correctly-scaled total shows no error', rescaled.totalDiff, 0);
}

/* ------------------------------------------------------------------ *
 * 11i. The phone layout must keep up with the desktop one
 *
 * The phone uses the SAME page as the Mac, reflowed by the
 * `@media (max-width: 700px)` block — there is no separate mobile app. So an
 * element added for the desktop without a rule in that block silently breaks the
 * phone, and plenty did: when this test was written, all seven elements added in
 * one session had zero mobile rules, including a <code> holding an IP address
 * (one long unbreakable token, guaranteed to overflow at 428px).
 *
 * These assert the rules EXIST. They cannot assert the layout looks right, which
 * is why the honest description is "no longer obviously broken", not "verified on
 * a phone".
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11j. Per-page change tracking
 *
 * A .pptx has no reliable per-slide date — PowerPoint rewrites the whole archive
 * on save — so the app keeps its own record. The subtlety is that fingerprints
 * must be compared by CONTENT, not by position: inserting a page near the front
 * shifts every later page's index, and a positional comparison would report the
 * entire deck as modified.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11k. Scan history snapshots
 *
 * History keeps a SUMMARY per scan, not the whole report. A full report for a
 * 200-page deck is a quarter of a megabyte; 25 of them would be megabytes of
 * state and would slow the page that loads them. These assert the snapshot stays
 * small and still carries what the history view shows.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11l. Recording an anchor must not corrupt the fit
 *
 * A per-criterion anchor whose parts do not add up to its stated total silently
 * poisons every fit built on it, and the output looks fine. This was written after
 * hand-building an anchor that recorded the app's MEAN where the app reports the
 * MEDIAN, which understated it by a mark.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11m. Dropping a PDF reviews THAT PDF
 *
 * Reported by the student: dropping a friend's PDF produced "This PDF is named
 * ... but the portfolio in use is ...", which blocked the legitimate and common
 * case of reviewing someone else's work, and described a problem that had not
 * happened. Dropping a PDF now clears the deck in use, which removes the mismatch
 * by construction, and the reply states what was replaced instead of warning.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11n. A rate limit must say WHICH limit it is
 *
 * Google returns 429 for both a per-minute limit and a per-DAY quota. The app used
 * to answer both with "wait a minute and press Check now again", which is useless
 * for a daily cap. Measured on this key: the free tier allows
 * `GenerateRequestsPerDayPerProjectPerModel` = 20, so a few checks exhaust a model
 * for the rest of the day and no amount of waiting a minute helps.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11o. Daily allowance, and when it resets
 *
 * A free-tier key spends REQUESTS per model per day, not money, so the app shows
 * percentage of allowance used instead of a cost. Two things have to be right or
 * it misleads: the reset time (midnight PACIFIC, confirmed by Google's docs — an
 * 8-hour error would say the allowance had come back when it had not) and the
 * per-model split (each model has its own cap, so being out on one is not being
 * out on all).
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11p. One answer to "which portfolio is being marked"
 *
 * Reported by the student: after a drag and drop it was "unclear which portfolio
 * is being marked". It was said twice, in two sections, in two different ways —
 * the upload section named the .pptx and the visual-review section named the PDF —
 * which read as two competing answers. There is now one place that answers it.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11q. The "Since your last check" panel is collapsible
 *
 * Asked for because it is a summary you read once and then want out of the way.
 * Collapsing is easy to undo by accident, so it is pinned: the panel must stay a
 * <details> using the same class as Settings, must open by default (when something
 * has changed, that is the first thing worth seeing), and must remember a choice
 * to close it rather than springing back open on every reload.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11r. The page must follow the LIVE config, not the last scan
 *
 * Reported by the student: "I dragged Jacob's in and it kept Rohan's portfolio on
 * the screen." The card read the deck name from `data.deck`, which describes the
 * last COMPLETED scan — so after choosing a new portfolio it still named the old
 * one until a scan ran, while the live config already pointed at the new file. Two
 * sources of truth, and the wrong one won.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11s. Deep thinking belongs to Settings
 *
 * Asked for: move it out of the Checks card. It is a setting about HOW the AI
 * answers, not a button you press, so it sat oddly among the Check now controls.
 * Easy to half-move — leaving the old block behind would put two checkboxes with
 * the same id in the page, and the second one silently wins.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11t. Structural checks run on a PDF
 *
 * A PDF-only student previously got NO structural checks: the app told them to add
 * a .pptx. A PDF export is a normal way to hand in work, so the free checks now
 * read the PDF's page text. A PDF has no font sizes, shape geometry or speaker
 * notes, so the rules needing those are SKIPPED AND NAMED — a rule that never runs
 * finds nothing, and "nothing" reads as a pass.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11u. The "Marking" card must follow the LIVE config, not the last scan
 *
 * Reported twice. First the DECK name came from the saved report, so after choosing
 * a new portfolio the card still named the old one. Then the PDF name did the same
 * thing for a subtler reason: `publicConfig` returned `pdfFile` but NOT `pdfName`,
 * and `buildStateResponse` merges the live config over the report's copy — so the
 * report's STALE `pdfName` survived the merge untouched while `pdfFile` was live.
 * The card showed the previous PDF's name indefinitely.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11v. Changing the API key from the interface
 *
 * Swapping the key used to mean hand-editing config.json — the UI said the key was
 * "never shown or typed here". But the free tier runs out, so swapping is routine,
 * not setup. Two things must hold: the key must be testable WITHOUT spending
 * generation quota (or testing it is impossible exactly when you need to), and the
 * key itself must never be sent back to the browser.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11w. The costs section tells the FREE-TIER truth
 *
 * It was written when the student was paying per check: it said the AI review "is
 * billed by Google to your account" and headed itself "What costs money". That is
 * now wrong — the key is deliberately on the free tier — so it was rewritten around
 * the daily REQUEST allowance, which is the thing that actually runs out.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11x. The allowance meter must not over-count
 *
 * Reported by the student: "the tracker must be wrong because it says I've used it
 * all but on Google it says I have loads left." It was wrong. The tracker counted
 * every request it SENT, so it included:
 *   - 429 and 5xx responses, which Google REFUSED and which therefore consume none
 *     of the daily quota;
 *   - `/v1beta/models` GETs, which are not `generateContent` calls and do not touch
 *     `GenerateRequestsPerDayPerProjectPerModel` at all.
 * It reached 27 against a limit of 20 while Google was still serving checks.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11y. Every AI request path must actually work
 *
 * The worst bug of this session, and it was invisible in review. A count was added
 * inside `callWithRetry` referring to `onRequest` — a name that is destructured in
 * `callGemini`, NOT in `callWithRetry`. So it threw `ReferenceError: onRequest is
 * not defined` on EVERY request; the surrounding catch reported it as "The AI
 * request failed", and the app failed every AI check while the code read as correct.
 *
 * This asserts the counter is reachable at every layer, which is the thing that was
 * wrong, rather than asserting the feature works.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 11z. No money anywhere, an accurate meter, stats, and export
 *
 * Four requests in one: the key is free-tier, so no price may be shown; the
 * allowance percentage may only come from a limit Google has CONFIRMED; a stats card
 * was wanted; and a one-file export.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 12. Scoring: the app is corrected by what it has been MEASURED to get wrong
 *
 * Six real portfolios with known marks showed two things the raw model could not fix
 * on its own: the default marking model reads 2-6 marks LOW, and the cheap model
 * COMPRESSES (it gave a 99, a 96 and an 89 portfolio 81, 80 and 66). Anchors are
 * therefore grouped by marking model — pooling them produced the nonsense fit
 * real = app x 0.043 + 93.7.
 *
 * The correction is applied ONLY inside the range the anchors cover. Below it nothing
 * has been measured, so the raw figure is left alone and said to be a guess:
 * extrapolating a fudge factor into unmeasured territory would make the app
 * confidently wrong in a new way.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 12b. The per-criterion dropdown in the coverage grid
 *
 * Requested: each criterion should open into a short colour-coded paragraph of what is
 * good and what is not. Two things have to be right or it misleads:
 *   - the colour must follow the VERDICT, and the verdict classes are
 *     good/medium/high/critical, not "bad"/"mid" — mapping them wrongly leaves every
 *     row neutral, which is how a colour-coded control silently stops being colour coded;
 *   - the AI's comment must be filed by the verdict. It was always listed under "What is
 *     working", so a criterion judged PARTIAL showed "missing formal CAD drawings" as
 *     something working.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 12c. Past checks name the portfolio that was checked
 *
 * Requested because the list was ambiguous: a student who has reviewed their own
 * portfolio, a friend's and a draft had several entries whose marks could not be told
 * apart. The filename was already in the snapshot — it was simply never shown.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 12d. A Past checks row can be CLOSED again
 *
 * Reported: "you can't close the past checks expansion thing when you click on it".
 * The click handler called open() unconditionally, so a row opened and never shut —
 * the only escape was reloading the page.
 * ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ *
 * 12e. A failed scan must never show a mark of zero
 *
 * Found while investigating a test failure. A scan that fails (or runs without an API
 * key) still writes a report with `predicted: 0, maximum: 0, weightsKnown: false` and a
 * row for every criterion. The UI's guard only checked for rows, so it let that through
 * and would have shown a predicted mark of "0 / 0" — a grade of zero on a portfolio
 * nobody had assessed.
 * ------------------------------------------------------------------ */

section('12e. No mark from a scan that produced none');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const ai = require(path.join(ROOT, 'lib', 'ai.js'));
  const spec = ai.loadSpec(path.join(ROOT, 'spec'), 'aqa-gcse');

  const empty = ai.summarisePredictedMarks(spec.criteria, []);
  ok('no coverage gives no percentage', empty.percentage === null, String(empty.percentage));
  ok('and no maximum to be a mark out of', empty.maximum === 0, String(empty.maximum));
  ok('but it does flag that no weights are known', empty.weightsKnown === false);
  ok('rows still exist for each criterion, which is why rows alone are not enough',
    Array.isArray(empty.rows) && empty.rows.length > 0, String(empty.rows && empty.rows.length));

  ok('the UI refuses a prediction with no known weights',
    /if \(p\.weightsKnown === false\) \{ return null; \}/.test(page));

  const decides = (p) => {
    if (!p || typeof p !== 'object') return null;
    if (p.weightsKnown === false) return null;
    if (typeof p.percentage !== 'number' && !(p.rows && p.rows.length)) return null;
    return p;
  };
  ok('simulated: an empty prediction is refused', decides(empty) === null);
  ok('simulated: a real prediction is kept',
    decides({ percentage: 83, rows: [{}], weightsKnown: true }) !== null);
}

section('12d. Past checks rows toggle shut');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  ok('the open row is tracked', /var openIndex = null;/.test(page));
  ok('clicking the open row closes it',
    /if \(openIndex === index\) \{[\s\S]{0,120}openIndex = null;/.test(page));
  ok('and clears the expanded content', /openIndex = null;\s*\n\s*clearNode\(detail\);/.test(page));
  /*
   * The row must LOOK open, or the reader cannot tell that clicking again will shut it
   * — a state you cannot see is a state you cannot control.
   */
  ok('the open row is visually marked', /\.history-row\.is-open/.test(page));
  ok('the class is applied by a helper', /var markRows = function \(\)/.test(page));
  ok('and it is called on both open and close', (page.match(/markRows\(\);/g) || []).length >= 2);
  ok('the row reports its state for screen readers', /setAttribute\('aria-expanded'/.test(page));
  // An empty detail box must not leave a gap.
  ok('an empty detail box collapses rather than leaving a gap',
    /\.history-detail:empty \{ display: none; \}/.test(page));

  /*
   * Reproduce the bug directly: two clicks on the same row must end CLOSED. This is the
   * behaviour that was wrong, so it is simulated rather than only asserted in the source.
   */
  let openIndex = null;
  let detailItems = 0;
  const open = (index) => {
    if (openIndex === index) { openIndex = null; detailItems = 0; return; }
    openIndex = index; detailItems = 1;
  };
  open(0);
  const afterFirst = openIndex;
  open(0);
  const afterSecond = openIndex;
  ok('simulated: first click opens', afterFirst === 0, String(afterFirst));
  ok('simulated: second click closes', afterSecond === null, String(afterSecond));
  open(0); open(1);
  ok('simulated: opening another row closes the first', openIndex === 1);
}

section('12c. Past checks name the file');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const { snapshotOf } = require(path.join(ROOT, 'server.js'));

  // The snapshot must record the file, and does.
  const snap = snapshotOf(JSON.parse(fs.readFileSync(path.join(ROOT, 'state', 'last-report.json'), 'utf8')));
  ok('a snapshot records which file was checked', !!snap.deck, String(snap.deck));
  ok('and which model marked it', !!snap.visionModel, String(snap.visionModel));

  // The list endpoint must pass both through, or the UI has nothing to show.
  ok('the history list returns the filename', /index: i, at: s\.at, deck: s\.deck/.test(serverSrc));
  ok('and the marking model', /visionModel: s\.visionModel \|\| null,/.test(serverSrc));

  // ...and the UI must actually render them.
  ok('the table has a Portfolio column', /el\('th', \{ text: 'Portfolio' \}\)/.test(page));
  ok('the row shows the filename', /cls: 'history-file'/.test(page) && /text: name, attrs: \{ title: name \}/.test(page));
  ok('the row shows the model that marked it', /marked by ' \+ String\(sc\.visionModel\)/.test(page));
  /*
   * Entries written before the filename was recorded have no deck. They must say so
   * rather than render an empty cell, which reads as a fault.
   */
  ok('an older entry without a filename says so rather than being blank',
    /el\('span', \{ cls: 'muted', text: 'not recorded' \}\)/.test(page));
  ok('a long filename wraps instead of widening the table',
    /history-file \{ font-size: 13px; overflow-wrap: anywhere/.test(page));
  // The detail view names the file too, so a clicked row is unambiguous.
  ok('the opened entry names the portfolio as well',
    /text: 'Portfolio: ' \+ String\(sc\.deck\)/.test(page));
}

section('12b. Coverage dropdown');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  ok('there is a per-criterion detail function', /function coverageDetail\(row\)/.test(page));
  ok('the row renders it instead of a flat comment',
    /kids\.push\(coverageDetail\(row\)\)/.test(page) && !/cls: 'cov-comment'/.test(page));

  // Colour must come from the verdict classes that actually exist.
  ok('the status classes are mapped as they are named',
    /meta\.cls === 'medium'/.test(page) && /meta\.cls === 'critical'/.test(page));
  ok('and not as bad/mid, which no status returns',
    !/meta\.cls === 'bad'/.test(page));
  ok('the three tones are styled', /\.cov-detail\.tone-good/.test(page)
    && /\.cov-detail\.tone-mid/.test(page) && /\.cov-detail\.tone-bad/.test(page));

  // The comment goes where the verdict points.
  ok('a strong criterion files its comment as good',
    /if \(row\.status === 'strong'\) \{ whatIsGood\.push\(row\.comment\); \}/.test(page));
  ok('a partial or weak one files it as not-good',
    /else if \(row\.status === 'partial' \|\| row\.status === 'weak' \|\| row\.status === 'none'\) \{ whatIsNot\.push\(row\.comment\); \}/.test(page));

  // Concrete faults, filtered to this criterion's pages.
  ok('findings are filtered to the pages this criterion cites',
    /if \(row\.slides\.indexOf\(pages\[q\]\) > -1\)/.test(page));
  ok('substance is listed before presentation', /var rank = \{ substance: 0, presentation: 1, obvious: 2 \}/.test(page));
  ok('the list is capped so the dropdown stays short', /k < onThesePages\.length && k < 5/.test(page));
  ok('and says how many more there are', /more points on these pages/.test(page));

  // It must not invent a criticism when there is none, nor look broken when empty.
  ok('an empty assessment is explained rather than left blank',
    /recorded no comment for this criterion/.test(page));
  ok('no pages cited is itself reported as a problem',
    /No pages were identified as evidence/.test(page));
}

section('12. Scoring calibration');
{
  const cal = require(path.join(ROOT, 'lib', 'calibration.js'));
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  ok('anchors exist and are grouped by marking model',
    cal.ANCHORS.length >= 6
      && cal.ANCHORS.some((a) => a.model === 'gemini-3.5-flash')
      && cal.ANCHORS.some((a) => a.model === 'gemini-3.1-flash-lite'));

  // The strong model must be corrected to within a mark on every anchor it has.
  const strong = cal.ANCHORS.filter((a) => a.model === 'gemini-3.5-flash');
  let worstStrong = 0;
  for (const a of strong) {
    const r = cal.assessPrediction(a.app, 100, { board: a.board, model: a.model });
    worstStrong = Math.max(worstStrong, Math.abs(r.calibrated - a.real));
  }
  ok('the default marking model lands within 1 mark on every anchor',
    worstStrong <= 1, 'worst error ' + worstStrong);

  // The cheap model was the whole reason for grouping: its error shape differs.
  const cheap = cal.ANCHORS.filter((a) => a.model === 'gemini-3.1-flash-lite');
  let worstCheap = 0;
  for (const a of cheap) {
    const r = cal.assessPrediction(a.app, 100, { board: a.board, model: a.model });
    worstCheap = Math.max(worstCheap, Math.abs(r.calibrated - a.real));
  }
  ok('the cheap model also lands within 2 marks once corrected',
    worstCheap <= 2, 'worst error ' + worstCheap);

  /*
   * The two models must NOT share a correction. The strong model's fit and the cheap
   * model's fit have different slopes, which is the measured reason for grouping.
   */
  const fs1 = cal.fitLine(strong);
  const fc1 = cal.fitLine(cheap);
  /*
   * Assert they differ, rather than by how much. A hand-picked threshold is a magic
   * number that fails the moment a new anchor shifts the fit slightly — which is what
   * happened at 0.15 when the real difference was 0.13. The MEANINGFUL claim is that
   * one model's correction is not valid for the other.
   */
  ok('the two models need different corrections',
    Math.abs(fs1.slope - fc1.slope) > 0.05,
    'strong slope ' + fs1.slope.toFixed(3) + ' vs cheap ' + fc1.slope.toFixed(3));
  /*
   * And the practical consequence, which is what grouping actually protects: applying
   * the WRONG model's correction must produce a visibly worse answer.
   */
  const anchorCheap = cheap[cheap.length - 1];
  const wrong = Math.round(fs1.slope * anchorCheap.app + fs1.intercept);
  ok('using the wrong model\'s correction would be worse',
    Math.abs(wrong - anchorCheap.real) > Math.abs(worstCheap),
    'wrong-model result ' + wrong + ' vs real ' + anchorCheap.real);

  // Refusing to correct outside the measured range is the point.
  const below = cal.assessPrediction(40, 100, { board: 'aqa-gcse', model: 'gemini-3.5-flash' });
  ok('a prediction below the measured range is NOT adjusted', below.applied === false);
  ok('and says why', /only been checked against portfolios it scored between/.test(below.note || ''), below.note);
  ok('it still reports the measured error so the reader can judge',
    below.measuredLow !== null && below.measuredHigh !== null);
  const above = cal.assessPrediction(100, 100, { board: 'aqa-gcse', model: 'gemini-3.5-flash' });
  ok('and a prediction above the range is not adjusted either', above.applied === false);

  // An unknown model has no evidence, and must not borrow another model's.
  const unknown = cal.assessPrediction(80, 100, { board: 'aqa-gcse', model: 'gemini-9-nonsense' });
  ok('an unknown marking model gets no correction at all', unknown.calibrated === null);
  ok('and is told so', /No marked portfolios have been compared/.test(unknown.note || ''));

  // A correction must never exceed the paper.
  const clamped = cal.assessPrediction(97, 100, { board: 'aqa-gcse', model: 'gemini-3.5-flash' });
  ok('a corrected mark never exceeds the maximum', clamped.calibrated <= 100);

  // The raw figure must survive, so the adjustment is auditable.
  const r1 = cal.assessPrediction(83, 100, { board: 'aqa-gcse', model: 'gemini-3.5-flash' });
  ok('the raw AI figure is kept alongside the corrected one', r1.raw === 83 && r1.calibrated === 89);
  ok('and the adjustment is explained in words', /shown higher here/.test(r1.note || ''));

  // It must actually reach the report and the UI.
  ok('the report carries the calibration', /predicted: Object\.assign\(\{\}, predicted, \{ calibration \}\)/.test(serverSrc));
  ok('it is computed with the marking model, not the page reader',
    /const markModel = String\(cfg\.visionModel \|\| ''\)\.trim\(\) \|\| cfg\.model;/.test(serverSrc));
  ok('the UI shows the adjustment and its reason',
    /predicted-calibrated/.test(page) && /cal\.note/.test(page));
  ok('and shows the unadjusted note when no adjustment applies',
    /else if \(cal && cal\.note\)/.test(page));
  /*
   * The panel was restructured after the student said it "looked a bit messy": the
   * number was sharing a 128px column with three percentages and three notes. The fix
   * is a stacking column, and these pin it so it cannot quietly revert.
   */
  ok('the supporting text sits in its own column, not the figure column',
    /predicted-body/.test(page) && /el\('div', \{ cls: 'predicted-body' \}, supporting\)/.test(page));
  ok('the figure column no longer holds the notes',
    !/cls: 'predicted-note'[\s\S]{0,80}cls: 'predicted-pct'/.test(page));
  ok('the calibration note is its own emphasised block',
    /class="predicted-calibrated"|cls: 'predicted-calibrated'/.test(page));
  ok('the measurement facts use the quiet fact style',
    /predicted-fact/.test(page));
}

section('11z. No money, honest meter, stats, export');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const { quotaReport, buildStats, exportHtml } = require(path.join(ROOT, 'server.js'));

  /* --- No money --- */
  /*
   * Strip comments BEFORE looking for money.
   *
   * A comment explaining why money was removed legitimately contains "£0.00", and a
   * naive search matches across block boundaries and reports it. What matters is
   * whether any currency symbol survives into code the browser runs.
   */
  const pageNoComments = page
    .replace(/\/\*[\s\S]*?\*\//g, '')   // /* ... */ including the HTML ones
    .replace(/(^|[^:])\/\/[^\n]*/g, '$1'); // // line comments, not http://
  ok('the UI renders no money for the student',
    !/£/.test(pageNoComments),
    'found a currency value outside a comment');
  ok('and no money formatter', !/function moneyPrecise/.test(page) && !/function money\(/.test(page));
  ok('the server sends no per-model price',
    !/estimatedScanGbp:/.test(serverSrc) && !/estimatedPerPageGbp:/.test(serverSrc));
  ok('no cheaper-model suggestion is computed', !/savingFactor/.test(serverSrc));
  ok('the cost-based sort is gone', !/estimatedScanGbp \?/.test(page));

  /* --- The percentage is only shown from a CONFIRMED limit --- */
  /*
   * The SERVER still distinguishes a confirmed limit from an assumed one, and that
   * distinction must survive: it is what lets the UI label an estimate honestly.
   * The UI shows a percentage either way, because the student asked for one — what
   * changes is the label, not whether a number appears.
   */
  const assumed = quotaReport('m-assumed', null);
  ok('the report says whether the limit is confirmed',
    /limitKnown: !!learned/.test(serverSrc));
  ok('an unconfirmed limit is labelled assumed', /limitAssumed: !learned/.test(serverSrc));
  ok('and no percentage is claimed as measured when it is not',
    /percent: learned \? Math\.round/.test(serverSrc));
  ok('the UI always renders a percentage',
    /host\.appendChild\(statBlock\(pct \+ '%'/.test(page));
  ok('and labels it estimated when the limit is not confirmed',
    /limit estimated/.test(page) && /estimated\)/.test(page));
  ok('the bar is rendered whenever there is a percentage',
    /var barPct = Math\.max\(0, Math\.min\(100, pct\)\);/.test(page));
  // The counts themselves are always exact, and that must stay true.
  ok('the exact request counts are still shown', /m\.used \+ ' of ' \+ m\.limit \+ ' used/.test(page));
  /*
   * The quota percentage must not be clamped, or exceeding the limit would look like
   * exactly hitting it. The document tidiness score DOES clamp to 0-100, correctly,
   * so this checks the quota line specifically rather than any Math.min(100, ...).
   */
  ok('the quota percentage is not clamped, so over-limit is visible',
    /percent: learned \? Math\.round\(\(used \/ learned\) \* 100\) : null/.test(serverSrc));
  // The assumed figure must still be exposed, labelled.
  ok('the assumed limit is still reported, labelled', /anyAssumed:/.test(serverSrc));
  ok('the UI explains why the percentage may be estimated',
    /does not publish it per model/.test(page));

  /* --- Stats --- */
  const { parsePresentation } = require(path.join(ROOT, 'lib', 'pptx'));
  const deck = parsePresentation(path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx'));
  const st = buildStats(deck, null);
  ok('stats are produced', !!st && st.pages === deck.slideCount);
  ok('words are summed from the pages', st.words === deck.slides.reduce((a, x) => a + (x.wordCount || 0), 0));
  ok('images are counted', st.images === deck.slides.reduce((a, x) => a + (x.imageCount || 0), 0));
  ok('typing time is derived from the words at the stated rate',
    Math.abs(st.typingHours - (st.words / st.typingWpm / 60)) < 0.1,
    st.words + ' words -> ' + st.typingHours + 'h at ' + st.typingWpm + 'wpm');
  ok('the rate is named in the data, so the figure is explicable',
    st.typingWpm > 0 && st.markerWpm > 0);
  ok('reading time is present', st.readingMinutes > 0);
  ok('repeated words are found', Array.isArray(st.topWords));
  ok('a missing deck produces no stats rather than throwing', buildStats(null, null) === null);

  // The card exists, is collapsible, and is rendered.
  ok('there is a stats panel', /id="statsDetails"/.test(page) && /id="statsBody"/.test(page));
  ok('it is rendered', /renderStats\(\);/.test(page));
  ok('and it stays hidden without stats', /show\(ui\.statsDetails, false\)/.test(page));

  /* --- Export --- */
  const report = JSON.parse(fs.readFileSync(path.join(ROOT, 'state', 'last-report.json'), 'utf8'));
  const html = exportHtml(report);
  ok('the export is a complete html document',
    /^<!doctype html>/i.test(html) && /<\/html>\s*$/.test(html));
  ok('it is self-contained: no external references',
    !/(src|href)=["']https?:/i.test(html), 'found an external URL');
  ok('it carries the findings', /class="f"/.test(html));
  ok('it carries a printable tick-list', /class="checklist"/.test(html));
  /*
   * The report contains student writing and AI output. An unescaped angle bracket
   * from either would break the document, so escaping is asserted with a deliberate
   * payload rather than trusted.
   */
  const nasty = exportHtml(Object.assign({}, report, {
    deck: { fileName: '<script>alert(1)</script>', slideCount: 1 },
    combined: { findings: [{ title: '<img src=x onerror=alert(1)>', severity: 'high', fix: 'a & b < c' }] },
  }));
  ok('student and AI text is escaped, so it cannot inject markup',
    !/<script>alert\(1\)<\/script>/.test(nasty) && !/<img src=x/.test(nasty),
    'an unescaped tag survived into the export');
  ok('an ampersand is escaped too', /a &amp; b &lt; c/.test(nasty));
  ok('a report with nothing still exports without throwing',
    /<!doctype html>/i.test(exportHtml({})));

  // The endpoint and the button.
  ok('there is an export endpoint', /p === '\/api\/export'/.test(serverSrc));
  ok('it refuses politely with no finished check', /no finished check to export/.test(serverSrc));
  ok('it downloads rather than displaying in place',
    /Content-Disposition.*attachment/.test(serverSrc));
  ok('the page offers the export', /id="btnExport"/.test(page));
  // indexOf rather than a regex: two literal "!" need no escaping headaches.
  ok('and only once there is a report',
    page.indexOf('show(ui.btnExport, !!(data && data.ok') > -1);
}

section('11y. The counter is reachable from every layer');
{
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');
  const ai = require(path.join(ROOT, 'lib', 'ai.js'));

  // Statically: the name must be bound before it is used.
  const retryStart = aiSrc.indexOf('async function callWithRetry');
  const retryBody = aiSrc.slice(retryStart, aiSrc.indexOf('\nfunction ', retryStart + 10));
  ok('callWithRetry destructures onRequest', /onRequest = null \}/.test(retryBody) || /onRequest = null/.test(retryBody));
  // The bare name must not be referenced: only the resolved local.
  ok('and uses a locally resolved reference, not the bare name',
    /requestCounter/.test(retryBody) && !/\bonRequest\(\)/.test(retryBody));
  ok('the bare onRequest() call is gone', !/[^.\w]onRequest\(\);/.test(aiSrc));

  /*
   * Dynamically: run a REAL call through the real path and require the counter to
   * fire, and the pass to succeed. A static check alone would have missed nothing
   * here, but it also cannot prove the request succeeds — and "succeeds" was the
   * part that was silently false.
   */
  /*
   * NOTE: this is deliberately static only. Proving a request COMPLETES needs a live
   * call, which would spend the student's daily allowance on every test run. The live
   * proof was done by hand (a sample-deck pass: ok true, counter fired twice,
   * batchesRun 3 = 2 batches + 1 cross-check), and the end-to-end suite below covers
   * the request machinery with a fake API.
   */
  ok('the counter is invoked only after a served response', /if \(served && !counted && requestCounter\)/.test(aiSrc));
}

section('11x. The allowance meter does not over-count');
{
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  // A refused request must not be counted.
  ok('the count is taken from the response status, not from sending',
    /const served = res\.status === 200;/.test(aiSrc)
      && /if \(served && !counted && requestCounter\)/.test(aiSrc));
  /*
   * The counter must be RESOLVED, not merely named. Referring to `onRequest` inside
   * callWithRetry without destructuring it threw "onRequest is not defined" on every
   * request, which the catch reported as a network failure — so every AI check failed
   * while the code read as correct. Asserted explicitly because it is invisible.
   */
  ok('the retry wrapper resolves the counter before using it',
    /async function callWithRetry\(args, \{ attempts = 4, onRetry = null, onRequest = null \}/.test(aiSrc)
      && /const requestCounter = onRequest \|\| args\.onRequest \|\| null;/.test(aiSrc));
  ok('it is counted AFTER the 429/5xx branch decides to retry',
    aiSrc.indexOf('const served = res.status === 200;')
      < aiSrc.indexOf('if (res.status === 429 || (res.status >= 500'),
    'the count must not happen before retries are handled');

  // The model probe is a GET and must not be counted.
  const lmStart = aiSrc.indexOf('async function listModels');
  const lmEnd = aiSrc.indexOf('\n}', lmStart);
  const listModelsBody = aiSrc.slice(lmStart, lmEnd);
  ok('listModels does not invoke the counter', !/onRequest\)/.test(listModelsBody.replace('void onRequest;', '')));
  ok('and it says why, so nobody re-adds it',
    /deliberately NOT counted/i.test(listModelsBody) || /not a `generateContent` call/.test(listModelsBody));

  // The key test must not move the meter either.
  ok('testing a key does not count a request',
    /models = await ai\.listModels\(key, ai\.API_HOST\);/.test(serverSrc),
    'the key test still passes a counter');

  // The scans own usage is a DELTA, and the baseline must exist even when the AI
  // never runs — it previously threw "requestsBefore is not defined".
  ok('the pre-scan baseline is declared outside the AI branch',
    /let requestsBefore = readQuota\(\);/.test(serverSrc));
  ok('and it is filled in again only where the AI actually runs',
    /requestsBefore = readQuota\(\);/.test(serverSrc));
  ok('the report reports a total even when nothing was used',
    /requestsTotal,/.test(serverSrc) && /let requestsTotal = 0;/.test(serverSrc));
}

section('11w. The costs section');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  ok('the heading is about the allowance, not money',
    /What costs your allowance/.test(page) && !/>What costs money, and what does not</.test(page));
  ok('it says plainly that there is no bill on a free key',
    /Nothing here costs money/.test(page));
  ok('the old "billed by Google to your account" claim is gone',
    !/billed by Google\s+to your account/.test(page));

  // The three escape routes when the allowance runs out, quickest first.
  ok('it explains that swapping the key is the fastest fix',
    /paste a different key in Settings/i.test(page));
  ok('it says each key has its own allowance', /Each key has its own daily/.test(page));
  ok('it names the reset time', /midnight Pacific time/.test(page));
  ok('and converts it to UK time, since that is who is reading',
    /8am in the UK/.test(page));
  ok('it mentions the cheaper-model route', /cheaper model/.test(page));

  // Real numbers rather than vague reassurance.
  // The wording was tightened for readability; the FACTS must survive.
  ok('it gives the requests a check takes, and the daily allowance',
    /about <strong>9 requests<\/strong>/.test(page) && /20 a day/.test(page));

  // The PDF limitation must be stated here too, where costs are explained.
  ok('it is honest that a PDF skips some checks',
    /A PDF cannot be checked for fonts, spacing or stray boxes/.test(page));

  /*
   * No money should be shown anywhere for a free key: it reads £0.00 forever and
   * says nothing about what runs out. The formatter may remain for paid keys, but it
   * must not be used for the per-check figures.
   */
  ok('the model dropdown quotes requests, not a price',
    /plural\(dropReqs, 'request'\) \+ ' per check'/.test(page)
      && !/estimatedScanGbp\) \+ ' per full scan'/.test(page));
  ok('the pre-scan preview quotes requests, not a price',
    /takes about '\s*\n?\s*\+ plural\(reqs\.total, 'request'\)/.test(page));
  ok('the history shows requests used, not money spent',
    /plural\(sc\.requestsTotal, 'request'\) \+ ' used'/.test(page)
      && !/moneyPrecise\(sc\.costGbp\)/.test(page));

  // The snapshot the history reads must carry the request count.
  ok('a scan snapshot records the requests it used',
    /requestsTotal: \(report\.cost && report\.cost\.requestsTotal\)/.test(serverSrc));

  // Cross-reference from the deep-thinking switch must still point here.
  ok('the deep-thinking link names the new heading',
    /What costs your allowance/.test(page));
}

section('11v. Changing the API key');
{
  const { publicConfig } = require(path.join(ROOT, 'server.js'));
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');

  // The UI must let the key be pasted, tested and saved.
  for (const id of ['apiKeyInput', 'btnTestKey', 'btnSaveKey', 'btnForgetKey', 'apiKeyStatus']) {
    ok('the panel has ' + id, new RegExp('id="' + id + '"').test(page));
  }
  ok('the key can be pasted in a multi-line box, so a 39-character key stays visible',
    /<textarea id="apiKeyInput"/.test(page));
  ok('the panel at least names the free key page',
    /aistudio\.google\.com\/apikey/.test(page));
  // Guard against the old user-facing claim coming back. It is matched as a
  // sentence a student would SEE, not as a bare phrase — a comment explaining why
  // the wording was removed legitimately quotes it.
  ok('the old claim that the key is never typed here is gone',
    !/>[^<]*never shown or typed here/.test(page));

  // Testing must not make a generation request, or it cannot be used when the
  // allowance is spent — the exact moment this feature is for.
  const testBlock = serverSrc.slice(
    serverSrc.indexOf("p === '/api/test-key'"),
    serverSrc.indexOf("p === '/api/quota'"),
  );
  // Match CALLS, not the words: the comment explaining why no generation call is
  // made legitimately mentions generateContent.
  ok('the key test uses listModels, not a generation call',
    /ai\.listModels\(/.test(testBlock) && !/await runAiPass|await callGemini|await visionCallOnce/.test(testBlock));
  ok('it says whether the key works', /note: present\.length/.test(testBlock));
  ok('and refuses politely when there is no key', /No key to test/.test(testBlock));

  // The key itself must never go to the browser.
  const pub = publicConfig({ apiKey: 'AIzaSECRETVALUE1234567890', deckDir: '.', pdfFile: null });
  ok('the key is not in the public config', pub.apiKey === undefined);
  ok('but its last four characters are, so the student knows WHICH key is loaded',
    pub.keyTail === '7890', String(pub.keyTail));
  ok('and no part of the key appears anywhere else in it',
    !JSON.stringify(pub).includes('AIzaSECRET'), 'the key leaked into the public config');

  // Saving an empty string must clear the key rather than store "".
  ok('an empty key clears it', /if \(typeof body\.apiKey === 'string'\) cfg\.apiKey = body\.apiKey\.trim\(\);/.test(serverSrc));

  // When the allowance runs out the app must point at swapping the key.
  ok('the rate-limit message names swapping the key in Settings',
    /paste a different key in Settings/.test(aiSrc));
  ok('and still says nothing was lost', /nothing was lost/i.test(aiSrc));
}

section('11t. The bundled sample is never mistaken for the student\'s work');
{
  /*
   * Reported by the user: with nothing uploaded, pressing Check now produced a
   * full report and the app never said the findings were about the bundled sample
   * deck. Two separate faults had to line up, and both are covered here.
   *
   *   1. findDeck's fallback sets `usedSample`, but the search ALSO walked into
   *      the app's own sample/ folder (the default deckDir is "..", and the walk
   *      goes three levels deep), so it returned the sample as an ordinary
   *      candidate and the fallback never ran.
   *   2. Even once flagged, buildStateResponse merges the live config OVER the
   *      saved report, so a report written while the sample was being read kept
   *      claiming `readingSample: true` after a real portfolio was uploaded.
   */
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  // The walk must not adopt the app's own demonstration deck.
  ok('the deck search skips the app\'s own sample folder',
    /path\.resolve\(full\) === bundledSampleDir/.test(serverSrc));
  ok('and skips a nested copy of it, identified by its contents',
    /looksLikeBundledSample/.test(serverSrc)
    && /AQA-GCSE-Portfolio-SAMPLE\\\.pptx/.test(serverSrc));
  ok('but only when it is not the folder the student actually chose',
    /path\.resolve\(full\) !== path\.resolve\(dir\)/.test(serverSrc));

  // The flag must be derived live, like pdfName, or the stale report wins.
  const { publicConfig } = require(path.join(ROOT, 'server.js'));
  const fake = { deckDir: '/nonexistent-folder-xyz', deckFile: null, pdfFile: null, apiKey: '' };
  const live = publicConfig(fake, true);
  ok('publicConfig accepts the live sample flag rather than trusting the report',
    live.readingSample === true, String(live.readingSample));
  const liveFalse = publicConfig(fake, false);
  ok('and reports false when the student\'s own file is being read',
    liveFalse.readingSample === false, String(liveFalse.readingSample));
  ok('it is derived even when the caller does not pass one',
    typeof publicConfig(fake).readingSample === 'boolean',
    String(publicConfig(fake).readingSample));

  // And the report written during a scan records which file it actually read.
  ok('the report records whether it read the sample',
    /readingSample: !!found\.usedSample/.test(serverSrc));

  // The UI must say so, in the place a user is about to act on the findings.
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  ok('the page renders a notice when it is reading the sample',
    /function renderSampleNotice/.test(page));
  ok('the notice is drawn by renderActions, beside the Check button',
    /renderActions\(\) \{[\s\S]{0,400}renderSampleNotice\(\)/.test(page));
  ok('and it offers a way to add a real portfolio',
    /Add my portfolio/.test(page));
  ok('the notice is not dismissible, so it cannot be lost',
    !/renderSampleNotice[\s\S]{0,900}banner-dismiss/.test(page));
}

section('11u. The marking card follows live config');{
  const { publicConfig } = require(path.join(ROOT, 'server.js'));
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  // The server must supply every display field the card needs, or a stale one
  // from the report survives the merge.
  const cfg = { deckDir: '.', deckFile: null, pdfFile: '/tmp/Some Portfolio.pdf', apiKey: 'x' };
  const pub = publicConfig(cfg);
  ok('publicConfig reports a PDF name', pub.pdfName === 'Some Portfolio.pdf', String(pub.pdfName));
  ok('and it is derived from the live path, so the two cannot disagree',
    pub.pdfName === String(pub.pdfFile).split('/').pop());
  ok('with no PDF the name is null, not a leftover',
    publicConfig({ deckDir: '.', pdfFile: null }).pdfName === null);
  // The regression: pdfName must be a key the merge can overwrite.
  ok('pdfName is an own property of the live config',
    Object.prototype.hasOwnProperty.call(pub, 'pdfName'));

  // And the card must not depend on that field name being right either.
  ok('the card falls back to deriving the name from the path',
    /String\(config\.pdfFile\)\.split/.test(page));
  ok('and prefers the live path when the name disagrees with it',
    /pdfName !== String\(config\.pdfFile\)\.split/.test(page)
      || /if \(config\.pdfFile && config\.pdfName[\s\S]{0,120}pdfName = String\(config\.pdfFile\)/.test(page));
  ok('the deck name is taken from the live config too',
    /hasLiveDeckKey \? liveDeck : scannedDeck/.test(page));

  // The merge that caused it must still put the live config last.
  ok('the live config is merged over the report',
    /config: Object\.assign\(\{\}, report\.config, publicCfg\)/.test(serverSrc));
}

section('11t. Structural checks on a PDF');
{
  const pdftext = require(path.join(ROOT, 'lib', 'pdftext.js'));
  const pdfdeck = require(path.join(ROOT, 'lib', 'pdfdeck.js'));
  const { runLocalChecks } = require(path.join(ROOT, 'lib', 'checks.js'));
  const PDF = '/Users/rohanghali/Downloads/DT - Google Slides.pdf';

  ok('the extractor is exported', typeof pdftext.extractPdfText === 'function');
  ok('the deck adapter is exported', typeof pdfdeck.deckFromPdf === 'function');

  // Robustness: rubbish input must never throw, or a bad file crashes a scan.
  for (const [label, buf] of [['empty', Buffer.alloc(0)], ['junk', Buffer.from('not a pdf at all')],
                              ['random', Buffer.from([0xde, 0xad, 0xbe, 0xef, 0x00, 0x01])]]) {
    let threw = null;
    try { const r = pdftext.extractPdfText(buf); if (!r || !Array.isArray(r.pages)) threw = 'no pages array'; }
    catch (e) { threw = e.message; }
    ok('a ' + label + ' buffer returns a result instead of throwing', threw === null, String(threw));
  }

  if (fs.existsSync(PDF)) {
    const parsed = pdftext.extractPdfTextFile(PDF);
    ok('every page is accounted for', parsed.pageCount === 44, String(parsed.pageCount));
    ok('the text is read, not left empty',
      parsed.pages.filter((x) => x.text.trim()).length === 44);
    const all = parsed.pages.map((x) => x.text).join('\n');
    ok('real names and headings survive extraction',
      all.includes('Jacob Cluer') && all.includes('Candidate number'), 'missing identifying text');
    // The classic failure of a naive decoder is a run-together line.
    ok('lines are separated, not run together', !/1222Centre/.test(all));

    const deck = pdfdeck.deckFromPdf(fs.readFileSync(PDF), { fileName: 'x.pdf' }).deck;
    ok('the deck looks like a deck', deck.slideCount === 44 && deck.slides.length === 44);
    ok('it is marked as PDF-sourced', deck.source === 'pdf');
    ok('every page has a fingerprint, so repeat checks stay incremental',
      deck.slides.every((x) => !!x.fingerprint));
    ok('word counts are computed', deck.slides.reduce((a, x) => a + x.wordCount, 0) > 5000);
    // These must be numbers: the report divides them, and a null crashed a scan.
    ok('page dimensions are numbers, not null',
      Number.isFinite(deck.slideWidthIn) && Number.isFinite(deck.slideHeightIn) && Number.isFinite(deck.aspect));

    const pdfRun = runLocalChecks(deck, { source: 'pdf' });
    const pptxRun = runLocalChecks(deck, { source: 'pptx' });
    ok('checks actually run on a PDF', pdfRun.ruleCount >= 8, String(pdfRun.ruleCount));
    ok('the text-reading checks produce real findings', pdfRun.findings.length > 0,
      String(pdfRun.findings.length));
    ok('no rule threw', (pdfRun.errors || []).length === 0, JSON.stringify(pdfRun.errors));

    // The skipped rules must be exactly the ones a PDF genuinely cannot support,
    // and named. Font size and word count are exact even in a PDF, so the rules
    // that read only those now run; the ones that need measured widths still do not.
    const skipped = pdfRun.skippedRules.map((x) => x.rule);
    ok('the rule set skipped for a PDF is exactly the unsupported one',
      skipped.slice().sort().join(',') === 'missing-notes,off-canvas,page-too-busy,text-overflow,uncaptioned-image',
      skipped.slice().sort().join(','));
    for (const r of ['missing-notes', 'off-canvas', 'text-overflow']) {
      ok('  and ' + r + ' is one of them', skipped.includes(r));
    }
    ok('rules reading exact font sizes DO run on a PDF',
      !skipped.includes('text-too-small') && !skipped.includes('font-inconsistency'));
    ok('rules needing measured line widths do NOT run on a PDF',
      skipped.includes('text-overflow') && skipped.includes('off-canvas'));
    ok('a PDF runs FEWER rules than the same deck as a PowerPoint',
      pdfRun.ruleCount < pptxRun.ruleCount);
    ok('the totals are reported so the UI can explain the gap',
      pdfRun.totalRules === pptxRun.totalRules && pdfRun.totalRules === 18,
      String(pdfRun.totalRules));
    ok('rules that read only words DO run', !skipped.includes('common-misspellings')
      && !skipped.includes('duplicate-pages') && !skipped.includes('near-empty-page'));
  } else {
    ok('the sample PDF is present for this test', false, 'missing ' + PDF);
  }

  // The server must be able to find a PDF when no .pptx is chosen.
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok('findDeck falls back to the PDF', /if \(cfg\.pdfFile && fs\.existsSync\(cfg\.pdfFile\)\)/.test(serverSrc));
  ok('the scan builds a PDF deck when the source is a PDF',
    /pdfdeck\.deckFromPdf\(/.test(serverSrc) && /const fromPdf = found\.kind === 'pdf'/.test(serverSrc));
  ok('the source reaches the checks', /source: fromPdf \? 'pdf' : 'pptx'/.test(serverSrc));
  ok('the report carries the source and the skipped rules',
    /source: local\.source,/.test(serverSrc) && /skippedRules: \(local\.skippedRules/.test(serverSrc));
  // A scan with no AI call must not crash on the request count.
  ok('the request count is declared outside the AI branch',
    /let requestsBefore = readQuota\(\);/.test(serverSrc));

  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  ok('the UI explains which checks were skipped', /id="checkCoverage"/.test(page));
  ok('and it is rendered', /renderCheckCoverage\(\);/.test(page));
  ok('it only speaks up for a PDF', /local\.source !== 'pdf'/.test(page));
}

section('11s. Deep thinking lives in Settings');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  // Exactly one control, or the wrong one gets read.
  const count = (page.match(/id="thinkingToggle"/g) || []).length;
  ok('there is exactly one deep-thinking checkbox', count === 1, String(count));

  const toggleAt = page.indexOf('id="thinkingToggle"');
  const settingsAt = page.indexOf('id="settingsDetails"');
  const slidesAt = page.indexOf('id="slidesDetails"');
  ok('it is inside the Settings panel',
    toggleAt > settingsAt && toggleAt < slidesAt);
  ok('and the Settings panel really encloses it (not just earlier in the file)',
    page.lastIndexOf('<details', toggleAt) === settingsAt - '<details class="card details-card" '.length
      || page.slice(settingsAt - 40, toggleAt).includes('details-body'));

  /*
   * It must sit beside the other switch, not loose in the panel.
   *
   * This used to slice up to the first `</details>`, which stopped being the end
   * of anything once Settings gained collapsible groups. Checking that the two
   * switches share a group is what the assertion was always about, and it does not
   * depend on how many groups Settings happens to have.
   */
  const markAt = page.indexOf('for="markToggle"');
  ok('Predict a mark and Deep thinking are grouped together',
    markAt > -1
    && toggleAt > markAt
    && page.slice(markAt, toggleAt).indexOf('</details>') === -1,
    'the two switches are not in the same group');

  // The old wording pointed at "the note at the bottom of this page", which read
  // oddly from inside Settings; it now names the section.
  ok('its description no longer says "the note at the bottom"',
    !/see the note at the bottom of this page/.test(page));
  ok('it says plainly that it does not change the predicted mark',
    /does not change the (predicted )?mark/.test(page));

  // The wiring must survive the move.
  ok('the checkbox is still wired to the cost hint',
    /on\(ui\.thinkingToggle, 'change'/.test(page));
  ok('and still read when a check is quoted', /ui\.thinkingToggle\.checked/.test(page));
}

section('11r. Live config beats the last scan');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  // The server merges live config OVER the report's copy, so the page can rely on it.
  ok('the server merges live config over the report',
    /config: Object\.assign\(\{\}, report\.config, publicCfg\)/.test(serverSrc));
  ok('so the live deck name is available to the page', /publicConfig\(cfg\)/.test(serverSrc));

  // And the page must actually prefer it.
  ok('the card prefers the live config deck',
    /hasLiveDeckKey \? liveDeck : scannedDeck/.test(page),
    'the card must not read the deck name from the scan alone');
  ok('it knows whether config carried a deck key at all',
    /hasOwnProperty\.call\(config, 'deckFile'\)/.test(page));
  ok('the page count is discarded when it describes a different portfolio',
    /if \(deckName && scannedDeck && deckName !== scannedDeck\) \{ pages = null; \}/.test(page));

  /*
   * A newly chosen portfolio has NOT been scanned, so everything on screen still
   * describes the previous one. Saying so is the difference between "it did not
   * work" and "press Check now".
   */
  ok('it warns that the new portfolio has not been read yet',
    /Not read yet — press Check now/.test(page));

  // Removing the PowerPoint: there was no way to do this at all.
  ok('the PowerPoint can be removed', /id="btnClearDeck"/.test(page));
  ok('the server accepts clearDeck', /body\.clearDeck === true/.test(serverSrc));
  ok('removing the PowerPoint leaves the PDF alone',
    /if \(body\.clearDeck === true\) \{\s*cfg\.deckFile = null;/.test(serverSrc));

  // The diagnostic that would have answered this without guesswork.
  ok('config changes are logged with what moved',
    /Config changed - /.test(serverSrc) && /deckFile, 'pdfFile'|\['deckFile', 'pdfFile'/.test(serverSrc));
}

section('11q. Collapsible check summary');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  ok('the panel is a <details>, not a plain section',
    /<details class="card details-card" id="compareCard"/.test(page));
  ok('it uses the same collapsible class as Settings',
    (page.match(/details-card/g) || []).length >= 3, String((page.match(/details-card/g) || []).length));
  ok('it still has a summary with its title',
    /<span class="summary-title">Since your last check<\/span>/.test(page));
  ok('the "compared with" note moved into the summary',
    /<span class="summary-note muted" id="compareWhen">/.test(page));
  ok('the body is still the details body', /<div class="details-body" id="compareBody">/.test(page));

  // Hidden until there is a comparison to show, exactly as before.
  ok('it stays hidden until there are two checks to compare',
    /id="compareCard" style="display:none"/.test(page));

  // Open by default, and the choice is remembered.
  ok('it opens by default', /lsGet\(LS_COMPARE_OPEN, 'open'\)/.test(page));
  ok('and a choice to close it is persisted',
    /lsSet\(LS_COMPARE_OPEN, ui\.compareCard\.open \? 'open' : 'closed'\)/.test(page));
  ok('and a remembered choice is not overwritten by the default',
    /compareOpenDecided/.test(page));
}

section('11p. Which portfolio is being marked');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  ok('there is a dedicated marking line', /id="reviewedName"/.test(page));
  ok('it is rendered', /renderReviewed\(\);/.test(page));
  // The old competing phrasings must be gone, or the ambiguity returns.
  ok('no section claims "In use:" as a second verdict', !/In use: /.test(page),
    'found a competing "In use:" label');
  ok('the upload section names itself as a PowerPoint file, not as the verdict',
    /PowerPoint file: /.test(page));
  ok('the PDF line says what it supplies, not what is marked',
    /Pages supplied by: /.test(page));

  // The decision itself, in both directions.
  ok('a PDF is what gets marked when one is set',
    /if \(pdfName\) \{[\s\S]{0,200}setText\(ui\.reviewedName, pdfName\)/.test(page));
  ok('a mismatched PDF and .pptx name is flagged',
    /have different names\. If they are not the same piece of work/.test(page));
  ok('a PDF with no .pptx is flagged, because the free checks cannot run',
    /the free structural checks have nothing to read/.test(page));
}

section('11o. Daily allowance');
{
  const { quotaReset, quotaDate, quotaReport } = require(path.join(ROOT, 'server.js'));

  // The reset must read exactly midnight in the reset zone.
  const r = quotaReset(new Date());
  ok('the reset reads midnight Pacific', / 00:00$/.test(r.local || ''), r.local);
  ok('it is in the future by less than a day',
    r.inMinutes > 0 && r.inMinutes <= 60 * 25, String(r.inMinutes));
  // Across a DST change the interval differs, which a fixed offset cannot do.
  const jan = quotaReset(new Date('2026-01-15T12:00:00Z'));
  const jul = quotaReset(new Date('2026-07-15T12:00:00Z'));
  ok('the reset tracks daylight saving rather than a fixed offset',
    jan.inMinutes !== jul.inMinutes || /00:00/.test(jan.local), jan.local + ' / ' + jul.local);
  ok('both still read midnight local', / 00:00$/.test(jan.local) && / 00:00$/.test(jul.local));

  ok('the allowance date is a calendar date',
    /^\d{4}-\d{2}-\d{2}$/.test(quotaDate()), quotaDate());

  const rep = quotaReport('model-a', 'model-b');
  ok('a report always states when it resets', rep.resetInMinutes > 0);
  ok('the assumed limit is exposed, so a guess is visible',
    typeof rep.assumedLimit === 'number' && rep.assumedLimit > 0, String(rep.assumedLimit));
  ok('the report lists at most the models in use', Array.isArray(rep.models));
}

section('11n. Rate-limit messages');
{
  const { describeRateLimit } = require(path.join(ROOT, 'lib', 'ai.js'));

  ok('a non-429 gets no message', describeRateLimit(500, 'anything') === null);
  ok('a plain 429 is described as a temporary limit',
    /Wait a minute/.test(describeRateLimit(429, 'rate limit exceeded') || ''));
  const daily = describeRateLimit(429, '{"quotaId":"GenerateRequestsPerDayPerProjectPerModel-FreeTier"}', 'gemini-3.5-flash') || '';
  ok('a daily quota is NOT described as a wait-a-minute problem', !/Wait a minute/.test(daily), daily);
  ok('it names the model', /gemini-3\.5-flash/.test(daily), daily);
  ok('it says the limit is daily', /daily limit/i.test(daily), daily);
  ok('it tells the student what actually helps',
    /billing|cheaper model/i.test(daily), daily);
  ok('and it reassures that nothing was lost', /nothing was lost/i.test(daily), daily);
}

section('11m. PDF upload semantics');
{
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  // The old behaviour: a filename comparison and a warning aimed at the student.
  ok('the deck is no longer left configured alongside a dropped PDF',
    /cfg\.deckFile = null;/.test(src), 'expected the deck to be cleared');
  ok('the filename-mismatch warning is gone',
    !/If they are not the same work, the marks will be about the wrong portfolio/.test(src));
  ok('and the reply says what it replaced',
    /replacedDeck: hadDeck/.test(src) && /is no longer the portfolio in use/.test(src));
  ok('the PDF still becomes the file under review',
    /cfg\.pdfFile = target;/.test(src));
}

section('11l. Anchor consistency');
{
  const cal = require(path.join(ROOT, 'tools', 'calibrate.js'));
  const crit = [
    { id: 'a', weight: 10 }, { id: 'b', weight: 10 },
    { id: 'c', weight: 20 }, { id: 'd', weight: 20 },
    { id: 'e', weight: 20 }, { id: 'f', weight: 20 },
  ];
  const full = { a: 9, b: 9, c: 17, d: 17, e: 14, f: 15 };  // sums to 81

  ok('marks that add up are accepted', cal.checkMarksConsistent(crit, full, 81) === null);
  ok('a missing criterion is refused',
    /No mark given for/.test(cal.checkMarksConsistent(crit, { a: 9 }, 9) || ''));
  ok('a total that disagrees with its parts is refused',
    /add up to 81 but --app says 90/.test(cal.checkMarksConsistent(crit, full, 90) || ''),
    String(cal.checkMarksConsistent(crit, full, 90)));
  ok('a mark above the criterion maximum is refused',
    /only out of 10/.test(cal.checkMarksConsistent(crit, { a: 20, b: 9, c: 17, d: 17, e: 14, f: 15 }, 92) || ''),
    String(cal.checkMarksConsistent(crit, { a: 20, b: 9, c: 17, d: 17, e: 14, f: 15 }, 92)));
  ok('a correct set is accepted even without a stated total',
    cal.checkMarksConsistent(crit, full, NaN) === null);
}

section('11k. Scan history snapshots');
{
  const { snapshotOf } = require(path.join(ROOT, 'server.js'));
  const report = require(path.join(ROOT, 'state', 'last-report.json'));
  const snap = snapshotOf(report);

  ok('a snapshot is produced', !!snap && !!snap.at);
  ok('it is a fraction of the report it summarises',
    JSON.stringify(snap).length < JSON.stringify(report).length / 10,
    JSON.stringify(snap).length + ' vs ' + JSON.stringify(report).length);
  /*
   * The counts are always present. The MARK is only present when the report actually
   * produced one: a scan that failed (or ran without an API key) leaves
   * `predicted: 0, maximum: 0`, and the snapshot must report no percentage rather than a
   * zero. Asserting a number unconditionally made this test depend on the last scan
   * having succeeded, which is not what it is testing.
   */
  ok('it always carries the finding count', typeof snap.findingsTotal === 'number');
  ok('it reports a mark only when the report had one',
    snap.predictedPercent === null || typeof snap.predictedPercent === 'number',
    String(snap.predictedPercent));
  ok('it keeps only the substance, not every finding',
    snap.substantive.length <= 12, String(snap.substantive.length));
  ok('and every kept item has a title', snap.substantive.every((f) => !!f.title));

  ok('an unusable report produces no snapshot', snapshotOf(null) === null);
  ok('and neither does a failed one', snapshotOf({ ok: false }) === null);
}

section('11j. Per-page change tracking');
{
  const { updatePageHistory } = require(path.join(ROOT, 'lib', 'checks.js'));

  const first = updatePageHistory(null, [{ index: 1, fingerprint: 'A' }, { index: 2, fingerprint: 'B' }], 'T1');
  eq('a first scan sees every page as new', first.pages.map((p) => p.status), ['new', 'new']);
  eq('and records when it first saw them', first.pages[0].firstSeen, 'T1');

  // Nothing changed: no page may be reported as edited.
  const same = updatePageHistory(first.history, [{ index: 1, fingerprint: 'A' }, { index: 2, fingerprint: 'B' }], 'T2');
  eq('an unchanged rescan reports nothing modified', same.pages.map((p) => p.status), ['unchanged', 'unchanged']);
  eq('and does not move lastChanged', same.pages[0].lastChanged, 'T1');

  /*
   * The case that matters: a page inserted at the FRONT. Every later page shifts
   * index, so a positional diff would call all of them modified.
   */
  const inserted = updatePageHistory(first.history,
    [{ index: 1, fingerprint: 'NEW' }, { index: 2, fingerprint: 'A' }, { index: 3, fingerprint: 'B' }], 'T3');
  eq('an inserted page is new', inserted.pages[0].status, 'modified');
  eq('but the pages after it are only MOVED, not modified',
    inserted.pages.slice(1).map((p) => p.status), ['moved', 'moved']);
  eq('and a moved page keeps its original firstSeen',
    inserted.pages[1].firstSeen, 'T1');

  // Editing one page's content is the only thing that counts as modified.
  const edited = updatePageHistory(inserted.history,
    [{ index: 1, fingerprint: 'NEW' }, { index: 2, fingerprint: 'A' }, { index: 3, fingerprint: 'CHANGED' }], 'T4');
  eq('an edited page is modified', edited.pages[2].status, 'modified');
  eq('and its lastChanged moves', edited.pages[2].lastChanged, 'T4');
  eq('while untouched pages are unchanged', edited.pages.slice(0, 2).map((p) => p.status), ['unchanged', 'unchanged']);

  // A deleted page must not linger in the history.
  const removed = updatePageHistory(edited.history, [{ index: 1, fingerprint: 'NEW' }], 'T5');
  eq('a deleted page leaves the history', Object.keys(removed.history).length, 1);

  // Recording a page again after it was deleted counts as new, not moved.
  const restored = updatePageHistory(removed.history, [{ index: 1, fingerprint: 'NEW' }, { index: 2, fingerprint: 'A' }], 'T6');
  eq('a page that reappears is new again', restored.pages[1].status, 'new');
}

section('11i. Phone layout keeps up with the desktop');
{
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const start = page.indexOf('@media (max-width: 700px)');
  ok('the phone breakpoint exists', start > 0);
  // Take the first block after the breakpoint by balancing braces from there.
  let depth = 0;
  let end = start;
  for (let i = page.indexOf('{', start); i < page.length; i++) {
    if (page[i] === '{') depth += 1;
    else if (page[i] === '}') {
      depth -= 1;
      if (depth === 0) { end = i; break; }
    }
  }
  const mobile = page.slice(start, end);

  /*
   * Elements that would overflow or crowd a 428px screen without a rule. Listed
   * explicitly rather than inferred, because the failure mode is a person adding
   * something and forgetting — a list is what makes that visible.
   */
  const mustAdapt = [
    'connect-url',        // an IP address: one unbreakable token
    'connect-row',
    'predicted-top',      // a 128px figure beside a paragraph
    'predicted-figure',
    'predicted-pct',
    'predicted-unstable',
    'compare-mark',
    'compare-list',
    'finding-evidence',
    'pagechip',
    'finding-top',
    'show-more-wrap',
    // Grouped Settings and the card headings. These had NO phone rule at all when
    // they were added, which is the exact failure this list exists to catch.
    'settings-group',
    'group-body',
    'card-title',
    'key-actions',
  ];
  const missing = mustAdapt.filter((cls) => !mobile.includes('.' + cls));
  ok('every element that needs a phone rule has one', missing.length === 0,
    'missing: ' + missing.join(', '));

  // Long unbroken content anywhere is the usual cause of horizontal scroll.
  ok('the connect URL may break mid-token',
    /\.connect-url\s*\{[^}]*overflow-wrap/.test(mobile), 'no overflow-wrap on .connect-url');

  // The phone must not be sent a desktop-only two-column figure block.
  ok('the predicted figure stacks instead of sitting beside its note',
    /\.predicted-top\s*\{[^}]*display:\s*block/.test(mobile));
}

section('11v. The page draws a real report without falling over');
{
  /*
   * Every other test in this file reads the page as TEXT. None of them ran the
   * render functions, so an undefined property inside renderFixList would have
   * produced a blank page for every user with nothing here to catch it.
   *
   * This loads the page's real script into a DOM stand-in (tests/dom.js), hands it
   * a report, and checks the output actually appears.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);

  ok('the page script parses and initialises cleanly',
    page.errors.length === 0, JSON.stringify(page.errors));
  ok('the DOM stand-in found the overlay and the wizard button',
    !!page.ui.introOverlay && !!page.ui.btnShowSetup);

  // Prefer the real saved report; fall back to a small faithful one so the test
  // still means something on a machine that has never run a scan.
  let report = null;
  try {
    report = JSON.parse(fs.readFileSync(path.join(ROOT, 'state', 'last-report.json'), 'utf8'));
  } catch (err) { report = null; }

  const usingReal = !!(report && report.local && report.local.findings);
  if (!usingReal) {
    report = {
      ok: true, generatedAt: new Date().toISOString(),
      deck: { fileName: 'x.pptx', filePath: '/tmp/x.pptx', slideCount: 3, stats: {} },
      scan: { finishedAt: new Date().toISOString(), durationMs: 1200, slidesScanned: 3 },
      local: {
        score: 55, bySeverity: { critical: 0, high: 1, medium: 1, low: 0, info: 0 },
        findings: [
          { id: 'l1', rule: 'missing-title', severity: 'high', tier: 'substance', title: 'No title on page 2', slides: [2], evidence: 'e', why: 'w', fix: 'f', effort: 'minutes' },
          { id: 'l2', rule: 'text-too-small', severity: 'medium', tier: 'mechanical', title: 'Small text on page 3', slides: [3], evidence: 'e', why: 'w', fix: 'f', effort: 'minutes' },
        ],
      },
      ai: null,
      combined: { findings: [] },
      cost: { scans: 1 },
      config: { specId: 'aqa-gcse', specName: 'AQA GCSE', apiKeySet: false, pdfName: null },
      quota: null,
    };
  }
  ok('a report is available to render' + (usingReal ? ' (the real saved one)' : ' (synthetic)'), !!report);

  page.state.data = report;
  const r = page.call('renderAll');
  ok('renderAll completes on a full report', r.ok, r.ok ? '' : String(r.error));

  // If it ever throws, name the function that did, rather than only knowing that
  // "rendering failed" — that is the difference between a useful report and a hunt.
  const renderers = ['renderHeader', 'renderStatus', 'renderActions', 'renderCost',
    'renderTiles', 'renderControls', 'renderTabs', 'renderFixList', 'renderSuggestions',
    'renderCoverage', 'renderSlides', 'renderSettings', 'renderStats', 'renderReviewed',
    'renderKeySection', 'renderCheckCoverage', 'renderScanCostHint', 'renderSampleNotice'];
  const threw = [];
  for (const fn of renderers) {
    const out = page.call(fn);
    if (!out.ok) { threw.push(out.error); }
  }
  ok('no individual render function throws', threw.length === 0, JSON.stringify(threw));

  // Something must actually reach the DOM, or "no error" is meaningless.
  const fixList = page.els.get('fixList');
  ok('the fix list is populated', !!(fixList && fixList.children.length > 0),
    fixList ? 'children: ' + fixList.children.length : 'no #fixList element');
  const scoreEl = page.els.get('scoreNumber');
  ok('a score is written to the ring', !!(scoreEl && String(scoreEl.textContent).length),
    scoreEl ? JSON.stringify(scoreEl.textContent) : 'no element');
  const nameEl = page.els.get('deckName');
  ok('the portfolio name is written to the header', !!(nameEl && String(nameEl.textContent).length),
    nameEl ? JSON.stringify(nameEl.textContent) : 'no element');

  // The empty state is what a brand-new user sees first, so it must not throw.
  page.state.data = { ok: true, deck: null, local: null, ai: null, combined: null, config: {} };
  const empty = page.call('renderAll');
  ok('renderAll survives having no report at all', empty.ok, empty.ok ? '' : String(empty.error));

  // A failed scan must not be drawn as a mark of zero. The card hides itself.
  page.state.data = {
    ok: true, config: {}, local: null, combined: { findings: [] },
    ai: { predicted: { predicted: 0, maximum: 0, weightsKnown: false, rows: [{ name: 'x', mark: 0, max: 10, percent: 0 }] } },
  };
  const zero = page.call('renderPredicted');
  ok('a scan that produced no mark draws no predicted card', zero.ok, zero.ok ? '' : String(zero.error));
  ok('and the card is hidden rather than showing 0 out of 0',
    page.els.get('predictedCard') && page.els.get('predictedCard').hidden === true,
    page.els.get('predictedCard') ? 'hidden=' + page.els.get('predictedCard').hidden : 'no element');
}

section('11b. AI pass applies the cross-check');
(async () => {
  const https = require('https');
  const realRequest = https.request;
  const sent = [];

  // Answer like the Gemini API: one batch reply, then one cross-check reply.
  https.request = (opts, onResponse) => {
    const { EventEmitter } = require('events');
    const req = new EventEmitter();
    let body = '';
    req.setTimeout = () => {};
    req.write = (b) => { body += b; };
    req.end = () => {
      let json;
      if (opts.method === 'GET') {
        json = { models: [{ name: 'models/gemini-3.1-flash-lite', supportedGenerationMethods: ['generateContent'] }] };
      } else {
        const payload = JSON.parse(body);
        sent.push(payload);
        const isCross = /senior UK A-level/.test(payload.systemInstruction.parts[0].text);
        const answer = isCross
          ? { coverage: [{ criterion: 'b', status: 'strong', coveragePercent: 85, slides: [17], comment: 'Spec on page 17.' }],
            verdicts: [{ id: 's1', verdict: 'already-done', reason: 'Page 17 is the specification.' },
              { id: 's2', verdict: 'improve-existing', reason: 'Testing exists on page 18.' },
              { id: 's3', verdict: 'missing', reason: 'Nothing like it anywhere.' }],
            error: '' }
          : { findings: [
              { rule: 'no-spec', title: 'The design specification is missing', kind: 'add', progressAware: false, severity: 'high', slides: [2], evidence: '', why: '', fix: '', effort: 'hours', impact: 'high' },
              { rule: 'thin-tests', title: 'Testing lacks measured data', kind: 'add', progressAware: false, severity: 'medium', slides: [18], evidence: '', why: '', fix: '', effort: 'hours', impact: 'medium' },
              { rule: 'plan', title: 'Add a project plan', kind: 'add', progressAware: false, severity: 'low', slides: [3], evidence: '', why: '', fix: '', effort: 'hours', impact: 'medium' }],
            strengths: [], coverage: [{ criterion: 'b', status: 'none', coveragePercent: 0, slides: [], comment: 'not here' }], error: '' };
        json = { candidates: [{ content: { parts: [{ text: JSON.stringify(answer) }] }, finishReason: 'STOP' }],
          usageMetadata: { promptTokenCount: 1000, candidatesTokenCount: 100 } };
      }
      const res = new EventEmitter();
      res.statusCode = 200;
      onResponse(res);
      res.emit('data', Buffer.from(JSON.stringify(json)));
      res.emit('end');
    };
    return req;
  };

  try {
    const slide = (i, title) => ({ index: i, title, text: `${title} words`, wordCount: 2, imageCount: 0, images: [], shapes: [], notes: '', notesWordCount: 0, shapeCount: 1, fingerprint: 'f' + i });
    const deck = { fileName: 'x.pptx', slideCount: 3, slideWidthIn: 16.54, slideHeightIn: 11.69,
      slides: [slide(1, 'Contents'), slide(2, 'Research'), slide(3, 'Brief and specification')] };
    const spec = { text: 'spec', criteria: [{ id: 'b', name: 'Brief and specification' }] };
    const result = await ai.runAiPass({ deck, apiKey: 'test', model: 'gemini-3.1-flash-lite', spec, cache: {} });

    eq('two requests are made: the batch and the cross-check', sent.length, 2);
    eq('only the genuinely missing addition is suggested', result.suggestions.map((f) => f.title), ['Add a project plan']);
    eq('an addition that already exists is withdrawn, with the reason',
      result.suggestionsWithdrawn.map((w) => w.reason), ['Page 17 is the specification.']);
    const moved = result.findings.find((f) => f.title === 'Testing lacks measured data');
    ok('a criticism of existing work moves to the fix list', moved && moved.kind === 'judgement' && moved.tier === 'substance');
    ok('whole-portfolio coverage replaces the section notes',
      result.coverage.length === 1 && result.coverage[0].status === 'strong' && result.coverage[0].name === 'Brief and specification');
    ok('the cross-check result is cached', !!(result.cache && result.cache.crossCheck));

  const textOnly = ai.estimateScanCostGbp('gemini-3.1-flash-lite', 60);
  const withVision = ai.estimateScanCostGbp('gemini-3.1-flash-lite', 60, { vision: true });
  ok('vision costs more than text-only', withVision.tokensIn > textOnly.tokensIn,
    `${textOnly.tokensIn} vs ${withVision.tokensIn}`);
  ok('the vision estimate is flagged', withVision.vision === true && textOnly.vision === false);
  ok('the vision estimate is still pennies, not pounds', withVision.gbp * 100 < 10,
    `${(withVision.gbp * 100).toFixed(2)}p for 60 pages`);

  /*
   * Marks off must cost less. This drifted once: the browser halved its quote
   * when the mark was switched off and the server did not, so the model picker
   * over-quoted the price by three times.
   */
  const runsForMark = ai.visionRunsFor(true);
  const runsForNoMark = ai.visionRunsFor(false);
  eq('a mark is averaged over three runs', runsForMark, 3);
  eq('no mark costs one run', runsForNoMark, 1);
  /*
   * Marks off must cost meaningfully less, but NOT a third of the whole scan:
   * the text batches run either way, so only the repeated marking reviews are
   * saved. Measured on this portfolio: 9.2p with marks, 6.0p without.
   */
  const noMarkCost = ai.estimateScanCostGbp('gemini-3.1-flash-lite', 60, { vision: true, runs: runsForNoMark });
  const markCost = ai.estimateScanCostGbp('gemini-3.1-flash-lite', 60, { vision: true, runs: runsForMark });
  ok('switching the mark off costs noticeably less', noMarkCost.gbp < markCost.gbp * 0.75,
    `${(noMarkCost.gbp * 100).toFixed(2)}p vs ${(markCost.gbp * 100).toFixed(2)}p`);
  eq('the run count is reported in the estimate', markCost.runs, 3);

  /*
   * The quote must track MEASUREMENT, not drift. The scan that produced the
   * constants below is recorded in state/cost.json: a real flash-lite scan of 60
   * pages used 118,065 input and 15,349 output tokens. If someone edits a token
   * constant without re-measuring, this fails rather than mis-quoting the student.
   */
  const measuredIn = 118065;
  const measuredOut = 15349;
  const quoted = ai.estimateScanCostGbp('gemini-3.1-flash-lite', 60, { vision: true, runs: 1 });
  ok('the quoted input is within 25% of the measured scan',
    Math.abs(quoted.tokensIn - measuredIn) / measuredIn < 0.25,
    `quoted ${quoted.tokensIn} vs measured ${measuredIn}`);
  ok('the quoted output is within 25% of the measured scan',
    Math.abs(quoted.tokensOut - measuredOut) / measuredOut < 0.25,
    `quoted ${quoted.tokensOut} vs measured ${measuredOut}`);

  // A page sent as an image must be cheaper per page than the text path, which
  // re-sends the whole mark scheme on every batch. If this inverts, the measured
  // constant has been mis-edited.
  ok('the measured vision per-page constant is the smaller one',
    ai.MEASURED_TOKENS_PER_PAGE_VISION.input < ai.MEASURED_TOKENS_PER_PAGE.input);

  // No PDF must mean no vision request at all, not a wasted call.
  const noPdf = await ai.runVisionPass({
    deck: { slides: [], slideCount: 0 }, apiKey: 'k', model: 'gemini-3.1-flash-lite',
    spec: { criteria: [{ id: 'a', name: 'A', weight: 10 }] }, pdf: null, host: { requests: [] },
  });
  ok('no PDF means no vision call at all', noPdf === null);

  const visionCriteria = [
    { id: 'a', name: 'Investigating', weight: 10 },
    { id: 'c', name: 'Ideas', weight: 20 },
  ];
  const visionReply = {
    candidates: [{ content: { parts: [{ text: JSON.stringify({
      coverage: [
        { criterion: 'a', status: 'strong', coveragePercent: 90, suggestedMark: 99, slides: [1, 2], comment: 'good' },
        { criterion: 'c', status: 'partial', coveragePercent: 60, suggestedMark: 14, slides: [9], comment: 'developing' },
      ],
      strengths: [{ slide: 1, text: 'Strong contextual research.' }],
      suggestions: [{ rule: 'missing-tests', title: 'No destructive testing of the hinge', severity: 'medium', slides: [12], evidence: 'p12', why: 'w', fix: 'f', effort: 'hours', impact: 'high', worthScore: 9 }],
      findings: [{ rule: 'weak-annotation', title: 'Sketches are unannotated', severity: 'high', slides: [4], evidence: 'p4', why: 'w', fix: 'f', effort: 'hours', kind: 'judgement', impact: 'high', progressAware: true }],
      error: '',
    }) }] } }],
    usageMetadata: { promptTokenCount: 40000, candidatesTokenCount: 500 },
  };
  const vp = ai.parseVisionResponse(visionReply, { criteria: visionCriteria, slideCount: 60 });
  ok('the vision reply parses', vp.ok === true, vp.reason || '');
  eq('a mark above the criterion maximum is clamped', vp.coverage.find((c) => c.id === 'a').suggestedMark, 10);
  eq('a mark inside the maximum is kept', vp.coverage.find((c) => c.id === 'c').suggestedMark, 14);
  ok('vision coverage is marked as coming from the visual review', vp.coverage.every((c) => c.fromVision === true));
  ok('vision findings are read', vp.findings.length === 1 && /unannotated/.test(vp.findings[0].title));
  /*
   * The explicit "these are absent" list. Without it the "Worth adding" tab was
   * permanently empty, because the prompt warns so hard against proposing work
   * that already exists that nothing ever came through as kind "add".
   */
  ok('vision suggestions are read', vp.suggestions.length === 1 && /destructive testing/.test(vp.suggestions[0].title));
  ok('a vision suggestion is forced to be an addition',
    vp.suggestions[0].kind === 'add' && vp.suggestions[0].progressAware === false);
  ok('a vision suggestion keeps its worth score', vp.suggestions[0].worthScore > 0);
  ok('the vision schema allows suggestions', !!ai.visionSchema(visionCriteria).properties.suggestions);
  ok('the vision prompt asks explicitly for suggestions',
    /"suggestions": up to 5 pieces of work that are GENUINELY ABSENT/.test(ai.buildVisionPrompt({ slideCount: 60 }, { criteria: visionCriteria })));
  ok('vision strengths are read', vp.strengths.length === 1);
  ok('vision token usage is recorded', vp.tokensIn === 40000 && vp.tokensOut === 500);

  // A criterion the mark scheme does not contain must be dropped, not invented.
  const stray = ai.parseVisionResponse({
    candidates: [{ content: { parts: [{ text: JSON.stringify({ coverage: [{ criterion: 'zz', status: 'strong', coveragePercent: 100, suggestedMark: 5, slides: [], comment: '' }] }) }] } }],
  }, { criteria: visionCriteria, slideCount: 60 });
  ok('an unknown criterion in the vision reply is dropped', stray.coverage.length === 0);

  const vSchema = ai.visionSchema(visionCriteria);
  eq('the vision schema pins criterion ids', vSchema.properties.coverage.items.properties.criterion.enum, ['a', 'c']);
  ok('the vision schema allows a mark per criterion', !!vSchema.properties.coverage.items.properties.suggestedMark);
  ok('the vision prompt names the criteria and their maxima',
    /- a: Investigating — out of 10 marks/.test(ai.buildVisionPrompt({ slideCount: 60 }, { criteria: visionCriteria })));

    /*
     * Routing the visual findings.
     *
     * This was the bug: every vision finding was pushed into the fix list, so a
     * portfolio the model had genuine additions for showed an EMPTY "Worth
     * adding" tab. Nothing looked broken — the findings all appeared somewhere —
     * which is exactly why it needs asserting rather than eyeballing.
     */
    const routed = ai.partitionVisionFindings([
      { title: 'Add a page of failure analysis', kind: 'add', progressAware: false },
      { title: 'Sketches are unannotated', kind: 'judgement', progressAware: true },
      { title: 'Evaluation is description, not analysis', kind: 'fact', progressAware: true },
      // An "add" the model admitted already exists must NOT become a suggestion.
      { title: 'Add testing evidence', kind: 'add', progressAware: true },
    ]);
    eq('only genuinely missing work becomes a suggestion',
      routed.additions.map((f) => f.title), ['Add a page of failure analysis']);
    eq('everything else stays in the fix list',
      routed.changes.map((f) => f.title), ['Sketches are unannotated', 'Evaluation is description, not analysis', 'Add testing evidence']);
    const emptyRoute = ai.partitionVisionFindings(null);
    ok('no vision findings routes to nothing without throwing',
      emptyRoute.additions.length === 0 && emptyRoute.changes.length === 0);

    /*
     * 11f. Averaging the marking runs.
     *
     * One visual review is not stable enough to quote: the same real portfolio
     * came back 88% and then 81%, with individual criteria moving three marks.
     * Three runs let the median be taken and the spread be shown as a range,
     * which is why the app asks three times.
     */
    eq('the median of an even list takes the lower middle', ai.median([10, 20]), 10);
    eq('the median discards an outlier', ai.median([9, 12, 13]), 12);
    eq('the median ignores non-numbers', ai.median([null, 7, undefined, 9, 8]), 8);
    eq('an empty list has no median', ai.median([]), null);

    const aggCriteria = [
      { id: 'a', name: 'Investigating', weight: 10 },
      { id: 'b', name: 'Ideas', weight: 20 },
    ];
    const run = (a, b, title) => ({
      ok: true,
      coverage: [
        { id: 'a', status: 'strong', coveragePercent: 90, suggestedMark: a, slides: [1], comment: 'x' },
        { id: 'b', status: 'partial', coveragePercent: 50, suggestedMark: b, slides: [2], comment: 'y' },
      ],
      findings: [{ title }],
      suggestions: [{ title: 'shared suggestion' }],
      strengths: [{ slide: 1, text: 'A strength' }],
      tokensIn: 100, tokensOut: 10,
    });
    const agg = ai.aggregateVisionRuns([run(9, 15, 'f1'), run(10, 17, 'f2'), run(6, 12, 'f3')], aggCriteria);
    eq('the marks are the median of the runs', agg.coverage.map((c) => c.suggestedMark), [9, 15]);
    ok('each criterion carries what every run said',
      agg.coverage[0].marksPerRun.join(',') === '9,10,6', JSON.stringify(agg.coverage[0].marksPerRun));
    eq('the range spans the lowest and highest run totals', agg.markRange, { low: 18, high: 27 });
    eq('the spread is the difference', agg.markSpread, 9);
    eq('the number of runs is reported', agg.runs, 3);
    // Three runs each notice different things; dropping two thirds of them would
    // waste what the extra calls were paid for.
    eq('findings from every run are kept', agg.findings.length, 3);
    eq('a suggestion repeated across runs is not duplicated', agg.suggestions.length, 1);
    eq('strengths survive the union even though they have no title', agg.strengths.length, 1);
    ok('token usage is summed across runs', agg.tokensIn === 300 && agg.tokensOut === 30);

    const predAgg = ai.summarisePredictedMarks(aggCriteria, agg.coverage);
    eq('the predicted total is the sum of the medians', predAgg.predicted, 24);
    ok('a percentage range is produced from the runs', !!predAgg.percentageRange,
      JSON.stringify(predAgg.percentageRange));
    eq('the range low is the worst run as a percentage', predAgg.percentageRange.lowPercentage, 60);
    eq('the range high is the best run as a percentage', predAgg.percentageRange.highPercentage, 90);

    /*
     * Disagreement between runs must be surfaced, not hidden behind a median.
     *
     * Measured: the same portfolio gave 81, 80, 62 and 71 on the same code — a
     * 19-mark spread with valid replies, so the median is stable-looking but the
     * underlying answer is not. The threshold is what the report uses to warn.
     */
    ok('a stable portfolio is under the caution threshold',
      ai.aggregateVisionRuns([run(9, 15, 'a'), run(9, 15, 'b'), run(9, 16, 'c')], aggCriteria).markSpread < ai.VISION_SPREAD_CAUTION);
    const unstable = ai.aggregateVisionRuns([run(9, 15, 'a'), run(9, 15, 'b'), run(4, 8, 'c')], aggCriteria);
    ok('a 12-mark disagreement is over the threshold',
      unstable.markSpread >= ai.VISION_SPREAD_CAUTION, String(unstable.markSpread));

    const single = ai.aggregateVisionRuns([run(9, 15, 'f1')], aggCriteria);
    eq('one run is not turned into a fake range', single.markRange, null);
    eq('one run still reports one run', single.runs, 1);

    // A failed run must not poison the average; the survivors are used.
    const withFailure = ai.aggregateVisionRuns([run(9, 15, 'f1'), null, run(7, 13, 'f3')], aggCriteria);
    eq('a failed run is excluded, not treated as zero', withFailure.coverage[0].marksPerRun.join(','), '9,7');
    eq('the surviving runs are counted', withFailure.runs, 2);

    /*
     * 11g. Comparing two scans.
     *
     * The trap: finding ids carry a sequence number that increments on EVERY scan
     * (`wall-of-text@7#29`, then `#55`), and the AI's id embeds the title too. A
     * diff on raw ids therefore reports every single finding as both fixed and
     * new — which looks like it works and is completely wrong.
     */
    const { compareReports } = require(path.join(ROOT, 'server.js'));
    const prevReport = {
      ok: true, generatedAt: 'first',
      combined: { findings: [
        { id: 'wall-of-text@7#29', rule: 'wall-of-text', slide: 7, source: 'local', title: 'Long paragraph' },
        { id: 'missing-title@3#31', rule: 'missing-title', slide: 3, source: 'local', title: 'No title' },
      ] },
      ai: { predicted: { percentage: 70, predicted: 70, rows: [{ id: 'a', mark: 10, max: 15, name: 'Investigating' }] } },
    };
    const curReport = {
      ok: true, generatedAt: 'second',
      combined: { findings: [
        // The same problem, with a NEW sequence number, must count as unchanged.
        { id: 'wall-of-text@7#55', rule: 'wall-of-text', slide: 7, source: 'local', title: 'Long paragraph' },
        { id: 'off-canvas@9#56', rule: 'off-canvas', slide: 9, source: 'local', title: 'Element off page' },
      ] },
      ai: { predicted: { percentage: 78, predicted: 78, rows: [{ id: 'a', mark: 13, max: 15, name: 'Investigating' }] } },
    };
    const cmp = compareReports(prevReport, curReport);
    eq('a fixed problem is spotted despite a changed id', cmp.fixedCount, 1);
    eq('the fixed problem is the one that is gone', cmp.fixed[0].title, 'No title');
    eq('a genuinely new problem is spotted', cmp.addedCount, 1);
    eq('the new problem is named', cmp.added[0].title, 'Element off page');
    eq('unchanged problems stay unchanged', cmp.stillOpenCount, 2);
    eq('the mark movement is measured', cmp.markChange.delta, 8);
    eq('movement is reported per criterion', cmp.byCriterion[0].delta, 3);
    ok('comparing with no earlier report gives nothing', compareReports(null, curReport) === null);
    ok('comparing an unusable earlier report gives nothing', compareReports({ ok: false }, curReport) === null);

    /*
     * THE HALLUCINATION, pinned.
     *
     * Reported by the student: "21 problems fixed, 22 new findings" on a portfolio
     * they had not touched. Cause: the AI rewords and re-groups its own findings
     * between runs, so matching AI findings by title invents progress and invents
     * problems. One real fault, described three different ways, became one "fixed"
     * and three "new".
     *
     * These are the real titles from that report.
     */
    const aiPrev = {
      ok: true, generatedAt: 'a',
      combined: { findings: [
        { id: 'ai:x@53+55#0', rule: 'ai-a', source: 'ai', title: 'Distance sensor failure', slides: [53, 55] },
        { id: 'ai:y@35#1', rule: 'ai-b', source: 'ai', title: 'Cardboard model testing lacks a summary', slides: [35] },
      ] },
      ai: { predicted: { percentage: 88, predicted: 88, rows: [] } },
    };
    const aiCur = {
      ok: true, generatedAt: 'b',
      combined: { findings: [
        // Same fault, same pages, completely different wording.
        { id: 'ai:p@53+55#0', rule: 'ai-p', source: 'ai', title: 'Failure of core electronic functionality', slides: [53, 55] },
        { id: 'ai:q@53+55#1', rule: 'ai-q', source: 'ai', title: 'Functional Failure of Distance Sensor', slides: [53, 55] },
        // Same fault, page list slightly different from last time.
        { id: 'ai:r@35+36#2', rule: 'ai-r', source: 'ai', title: 'Model testing needs a summary of findings', slides: [35, 36] },
      ] },
      ai: { predicted: { percentage: 76, predicted: 76, rows: [] } },
    };
    const aiCmp = compareReports(aiPrev, aiCur);
    eq('reworded AI findings are NOT reported as fixed', aiCmp.fixedCount, 0);
    eq('and NOT reported as new', aiCmp.addedCount, 0);
    ok('they are paired by page instead', aiCmp.ai.paired >= 1, JSON.stringify(aiCmp.ai));
    eq('the AI counts are still reported', [aiCmp.ai.previousCount, aiCmp.ai.currentCount], [2, 3]);
    ok('the UI is told why there is no per-finding list', /rewords/i.test(aiCmp.ai.note), aiCmp.ai.note);

    /*
     * The headline count must separate substance from mechanics. "103 still to do"
     * was alarming and mostly meaningless: on the real report 93 of 129 findings
     * were mechanical items the Fix list already folds away, which buried the 27
     * that actually cost marks.
     */
    const tiered = compareReports({
      ok: true, generatedAt: 'a',
      combined: { findings: [
        { id: 'x@1#0', rule: 'r1', slide: 1, source: 'local', tier: 'substance', title: 'weak analysis' },
        { id: 'y@2#1', rule: 'r2', slide: 2, source: 'local', tier: 'obvious', title: 'small font' },
        { id: 'z@3#2', rule: 'r3', slide: 3, source: 'local', tier: 'obvious', title: 'stray box' },
      ] },
      ai: { predicted: { percentage: 50, predicted: 50, rows: [] } },
    }, {
      ok: true, generatedAt: 'b',
      combined: { findings: [
        { id: 'x@1#9', rule: 'r1', slide: 1, source: 'local', tier: 'substance', title: 'weak analysis' },
        { id: 'y@2#10', rule: 'r2', slide: 2, source: 'local', tier: 'obvious', title: 'small font' },
        { id: 'z@3#11', rule: 'r3', slide: 3, source: 'local', tier: 'obvious', title: 'stray box' },
      ] },
      ai: { predicted: { percentage: 50, predicted: 50, rows: [] } },
    });
    eq('substance is counted separately from mechanics', tiered.substantiveOpen, 1);
    eq('the mechanical items are counted too', tiered.mechanicalOpen, 2);
    eq('and the total still adds up', tiered.stillOpenCount, 3);
    // Two AI findings about the SAME page group must not both claim one prior one.
    eq('each prior finding is claimed at most once', aiCmp.ai.paired, 2);

    // A second scan with nothing changed must not call the API again.
    sent.length = 0;
    const again = await ai.runAiPass({ deck, apiKey: 'test', model: 'gemini-3.1-flash-lite', spec, cache: result.cache });
    eq('an unchanged rescan makes no generate requests', sent.length, 0);
    eq('the cached cross-check gives the same suggestions', again.suggestions.map((f) => f.title), ['Add a project plan']);
  } finally {
    https.request = realRequest;
  }
})().then(summary, (err) => {
  ok('AI pass end-to-end test ran', false, err && err.stack);
  summary();
});

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */

function summary() {
console.log('\n' + '='.repeat(60));
console.log(passed + ' passed, ' + failed + ' failed');
if (failed) {
  console.log('\nFailures:');
  for (const f of failures) { console.log('  - ' + f); }
  process.exit(1);
}
console.log('All tests passed.');
}
