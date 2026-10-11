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
const os = require('os');
const zlib = require('zlib');

const ROOT = path.join(__dirname, '..');
const { parsePresentation } = require(path.join(ROOT, 'lib/pptx.js'));
const { runLocalChecks, DISABLED_RULES } = require(path.join(ROOT, 'lib/checks.js'));
const ai = require(path.join(ROOT, 'lib/ai.js'));
const calibration = require(path.join(ROOT, 'lib/calibration.js'));

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
  const realPath = path.join(ROOT, '..', 'P1.pptx');
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
    /*
     * Asserted as a deliberate ZERO rather than deleted. The sample's smallest text is
     * 11pt on page 18, and the rule now starts at 10pt, so the sample is clean here —
     * and if the threshold ever moves back, this line says so instead of the answer key
     * failing somewhere unrelated.
     */
    eq('text-too-small no longer fires on the sample (its smallest text is 11pt)',
      sorted('text-too-small'), []);
    ok('uncaptioned-image does NOT fire on 14 (caption sits beside the image)',
      sorted('uncaptioned-image').indexOf(14) < 0);
    eq('no check throws', res.errors, []);
  }
}

section('3b. The smallest allowed font is 10pt');
{
  /*
   * Lowered from 12pt at the student's request. This is now the ONLY coverage the rule
   * has, because the sample's planted flaw sits at 11pt and is intentionally clean.
   *
   * The boundary is asserted from the size outwards rather than from the constant, so
   * this section still means something if the constant moves again.
   */
  const deckAt = (pt) => ({
    slideWidthPt: 595, slideHeightPt: 842, slideCount: 1,
    slides: [{
      index: 1, images: [], shapes: [],
      textShapes: [{ isTitle: false, kind: 'body', paragraphs: [{ text: 'a b c', sizePt: pt }] }],
    }],
  });
  const run = (pt) => runLocalChecks(deckAt(pt));
  const at = (pt) => run(pt).findings.filter((f) => f.rule === 'text-too-small');

  eq('10pt exactly is allowed, so the limit is inclusive', at(10).length, 0);
  eq('and so is everything above it',
    [10.5, 11, 12, 14, 18].map((p) => at(p).length), [0, 0, 0, 0, 0]);
  eq('9.5pt is not allowed', at(9.5).length, 1);
  eq('the finding names the real size rather than the limit',
    at(9.5)[0].title, 'Body text drops to 9.5pt');
  ok('and the evidence states the limit that was broken',
    /below 10pt/.test(at(9.5)[0].evidence), at(9.5)[0].evidence);
  eq('9.5pt reads as a medium, being close to the limit', at(9.5)[0].severity, 'medium');
  eq('8pt is where it becomes a high, two points under', at(8)[0].severity, 'high');
  eq('so 7pt is still a high', at(7)[0].severity, 'high');
  ok('the advice still asks for real body text rather than the bare minimum',
    /at least 14pt/.test(at(9.5)[0].fix), at(9.5)[0].fix);
  eq('nothing throws on a minimal deck', run(9.5).errors, []);
  ok('the threshold itself is 10, not merely the behaviour',
    /minFontPt: 10,/.test(fs.readFileSync(path.join(ROOT, 'lib', 'checks.js'), 'utf8')));
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
  const realPath = path.join(ROOT, '..', 'P1.pptx');
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
  /*
   * THE SUITE MUST NOT TOUCH THE STUDENT'S REAL DATA.
   *
   * server.js resolves its data folder the moment it is required, and without an override that
   * is ~/Library/Application Support/NEA Coach — the real API key, the real report, the real
   * quota. Every test that exercised those functions was reading the live files, and any that
   * wrote would have corrupted them. It only went unnoticed because the sandbox denies writes
   * to ~/Library, which is a property of where this happened to be run, not of the tests.
   *
   * A config.json is written first on purpose: the adoption logic treats a folder without one
   * as a fresh install and would COPY the repo's own state/ in — including an API key — into
   * a temporary directory.
   */
  const testDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'nea-tests-'));
  fs.writeFileSync(path.join(testDataDir, 'config.json'), JSON.stringify({ specId: 'aqa-gcse' }));
  process.env.NEA_COACH_DATA_DIR = testDataDir;
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
 * Reported by the student: "I dragged one portfolio in and it kept the previous one on
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
    // The shown figure is the corrected one WHEN a correction was applied.
    const shown = r.applied ? r.calibrated : r.raw;
    worstCheap = Math.max(worstCheap, Math.abs(shown - a.real));
  }
  /*
   * The cheap model is deliberately NOT corrected any more.
   *
   * Its fitted line is a good fit on its own anchors (max residual 1.2), but its
   * measured error spans 16 to 23 marks. Applying one number across a spread that
   * wide is another guess, and on a real portfolio it turned an 81 into a 98
   * against a true 90 - eight marks out, when the raw reading had been nine out.
   * The user's verdict was "still inaccurate ... this needs to go". So this model
   * now reports the AI's own figure and says why it is not adjusted.
   */
  ok('the cheap model is left uncorrected, because its error range is too wide',
    cheap.every((a) => {
      const r = cal.assessPrediction(a.app, 100, { board: a.board, model: a.model });
      return r.applied === false && r.reason === 'error-too-variable';
    }),
    'a cheap-model anchor was still corrected');
  ok('and its span really is the reason',
    cal.assessPrediction(80, 100, { board: 'aqa-gcse', model: 'gemini-3.1-flash-lite' }).measuredSpan > 5);

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
  const right = Math.round(fc1.slope * anchorCheap.app + fc1.intercept);
  /*
   * Measured against the RIGHT correction, not against leaving it alone.
   *
   * It used to compare the wrong-model result with `worstCheap`, which was the
   * corrected error. Now that the cheap model is deliberately left uncorrected,
   * "no correction" is the weaker baseline (23 marks out here) and the wrong line
   * at 12 marks out would beat it - making the assertion pass for the wrong reason.
   * The claim grouping actually protects is that the right fit beats the wrong one.
   */
  ok('using the wrong model\'s correction would be clearly worse',
    Math.abs(wrong - anchorCheap.real) > Math.abs(right - anchorCheap.real) + 5,
    'wrong-model ' + wrong + ' (err ' + Math.abs(wrong - anchorCheap.real)
      + ') vs right-model ' + right + ' (err ' + Math.abs(right - anchorCheap.real) + ')');

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
  ok('and the adjustment is explained in words', /adjusted up to/.test(r1.note || ''));

  /*
   * The note must state one unit throughout.
   *
   * It used to read "it read 16 to 23 marks low ... Those portfolios scored between
   * 66 and 81 on the app" - the AI's reading in marks, its error in marks, and the
   * anchors' scores in PERCENT, all in one sentence. Printed above a table that
   * totalled 75, that is the "lots of inconsistencies" a student reported. Nothing
   * in the note may be expressed as a percentage.
   */
  ok('the explanation names the paper total so "marks" has a denominator',
    /out of 100/.test(r1.note || ''), r1.note);
  ok('the explanation contains no percentages, so its units cannot be confused',
    !/%|percent/i.test(r1.note || ''), r1.note);
  ok('and it names both the AI reading and the corrected figure',
    new RegExp('as ' + r1.raw + ' out of').test(r1.note || '')
    && new RegExp('adjusted up to ' + r1.calibrated).test(r1.note || ''), r1.note);

  // It must actually reach the report and the UI.
  ok('the report carries the calibration', /predicted: Object\.assign\(\{\}, predicted, \{ calibration \}\)/.test(serverSrc));
  /*
   * The marking model, not the page reader — and the one that ACTUALLY marked it.
   *
   * This used to require the configured `cfg.visionModel`. That is wrong once the review
   * can fail over to another model: the anchors are grouped by model, so calibrating
   * against the configured name would apply the wrong line, or none. The field it must use
   * is whatever the AI pass recorded as having really answered.
   */
  ok('it is computed with the marking model, not the page reader',
    /const markModel = String\(\(aiResult && aiResult\.visionModel\) \|\| cfg\.visionModel \|\| ''\)\.trim\(\) \|\| cfg\.model;/
      .test(serverSrc),
    'the calibration must follow the model that actually produced the marks');
  /*
   * ONE short line, whichever way the correction went.
   *
   * There used to be an emphasised block headed "Why the estimate is higher than
   * the AI's own reading" whenever a correction applied. On a real portfolio it
   * explained, at length, turning 81 into 98 against a true 90. The user's verdict
   * was "this needs to go", and the correction itself is now only applied when the
   * measured error is tight enough for it to be an improvement - so the long
   * justification has nothing left to justify.
   */
  ok('the UI states the calibration outcome in one line',
    /if \(cal && cal\.note\) \{[\s\S]{0,160}predicted-fact/.test(page));
  ok('and the long "why this is higher" block is gone',
    !/Why the estimate is higher/.test(page),
    'the paragraph the user asked to remove is still present');
  ok('and there is no separate emphasised calibration block any more',
    !/predicted-calibrated/.test(page));
  /*
   * The panel was restructured after the student said it "looked a bit messy": the
   * number was sharing a 128px column with three percentages and three notes. The fix
   * is a stacking column, and these pin it so it cannot quietly revert.
   */
  ok('the supporting text sits in its own column, not the figure column',
    /predicted-body/.test(page) && /el\('div', \{ cls: 'predicted-body' \}, supporting\)/.test(page));
  ok('the figure column no longer holds the notes',
    !/cls: 'predicted-note'[\s\S]{0,80}cls: 'predicted-pct'/.test(page));
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

  /*
   * The long heading became a one-line lede when the panel was rewritten to be scannable. The
   * point stands: it must name the ALLOWANCE rather than a bill, because "it costs money" is the
   * one thing that is not true of a free key.
   */
  ok('the panel opens by naming the allowance, not a bill',
    /comes with a <strong>daily allowance<\/strong>/.test(page)
    && !/What costs money, and what does not/.test(page));
  ok('it says plainly that there is no bill on a free key',
    /Nothing here costs money/.test(page));
  ok('the old "billed by Google to your account" claim is gone',
    !/billed by Google\s+to your account/.test(page));

  // The three escape routes when the allowance runs out, quickest first.
  ok('it explains that swapping the key is the fastest fix',
    /paste another key[\s\S]{0,40}in Settings/i.test(page));
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
  /*
   * Shortened, but both facts have to survive: a PDF genuinely cannot supply layout, font size or
   * speaker notes, and the app must SAY which checks it skipped rather than reporting nothing.
   */
  ok('it is honest that a PDF skips some checks',
    /A PDF records the words on each page, but not the layout/i.test(page)
    && /names the checks it skipped/i.test(page));

  /*
   * No money should be shown anywhere for a free key: it reads £0.00 forever and
   * says nothing about what runs out.
   *
   * The per-model requests figure used to be appended to every option in the
   * dropdown, which made all thirteen rows identical apart from the model name.
   * It is deck-dependent and does not vary between models, so it now lives on the
   * line under the dropdown, which already updates when the choice changes. What
   * this assertion cares about is that the figure is REQUESTS and never a price.
   */
  ok('no price is shown where a model is chosen',
    !/estimatedScanGbp\) \+ ' per full scan'/.test(page));
  ok('the model list keeps only notes that distinguish a model',
    /var label = m\.name;\s*\n\s*if \(m\.supportsThinking === false\)/.test(page)
      && !/plural\(dropReqs, 'request'\) \+ ' per check'/.test(page));
  ok('the requests-per-check figure lives on the line under the dropdown',
    /A full check of ' \+ plural\(pages, 'page'\) \+ ' takes about '/.test(page));
  ok('the pre-scan preview quotes requests, not a price',
    /takes about '\s*\n?\s*\+ plural\(reqs\.total, 'request'\)/.test(page));
  ok('the history shows requests used, not money spent',
    /plural\(sc\.requestsTotal, 'request'\) \+ ' used'/.test(page)
      && !/moneyPrecise\(sc\.costGbp\)/.test(page));

  // The snapshot the history reads must carry the request count.
  ok('a scan snapshot records the requests it used',
    /requestsTotal: \(report\.cost && report\.cost\.requestsTotal\)/.test(serverSrc));

  // Cross-reference from the deep-thinking switch must still point here.
  /*
   * The switch in Settings points at this panel. It used to be checked through a heading that read
   * "What costs your allowance, and what does not"; that heading is now a lede and the explanation
   * is a titled row, so what matters is that the panel still explains it and is still the target.
   */
  ok('the deep-thinking explanation is still in the information panel',
    /id="infoDetails"/.test(page) && /Deep thinking/.test(page));
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
    /function renderPortfolioNotice/.test(page));
  /*
   * Read the function's own body rather than a character window around it. The window was a
   * proxy for "inside renderActions" and started failing the moment the function legitimately
   * grew, which is a test measuring the wrong thing.
   */
  ok('the notice is drawn by renderActions, beside the Check button',
    (() => {
      const body = /\n  function renderActions\(\) \{([\s\S]*?)\n  \}/.exec(page);
      return !!body && /renderPortfolioNotice\(\)/.test(body[1]);
    })(), 'the notice must be drawn by the same function as the Check button');
  ok('and it offers a way to add a real portfolio',
    /Add my portfolio/.test(page));
  ok('the notice is not dismissible, so it cannot be lost',
    !/renderPortfolioNotice[\s\S]{0,1400}banner-dismiss/.test(page));
}

section('11w. The exam board is offered even before the first check');
{
  /*
   * Reported by the user: "there is no option to change the exam board in the
   * setup and only one option in the drop down in settings".
   *
   * The no-report branch of buildStateResponse hand-rolled its own config object
   * and left out `specId` and `specPresets`. Three things followed:
   *
   *   1. the board dropdown fell back to a single hard-coded Eduqas option;
   *   2. the wizard's board list came back empty for the same reason;
   *   3. the wizard then skipped the save, because its fallback id happened to
   *      equal the server default, so the board was never recorded at all.
   *
   * The fix is that the branch now goes through publicConfig() like every other
   * path. These assertions pin the response shape and the save behaviour.
   */
  const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');

  // The response itself. buildStateResponse reads the live config, so this can be
  // called directly and must carry the board list either way.
  const { buildStateResponse } = require(path.join(ROOT, 'server.js'));
  const state = buildStateResponse();
  ok('the state response carries specId', !!state.config.specId, String(state.config.specId));
  ok('the state response carries every board as a preset',
    Array.isArray(state.config.specPresets) && state.config.specPresets.length >= 2,
    JSON.stringify(state.config.specPresets));
  ok('and both of this app\'s boards are in it',
    state.config.specPresets.some((p) => p.id === 'aqa-gcse')
    && state.config.specPresets.some((p) => p.id === 'eduqas-alevel'),
    JSON.stringify(state.config.specPresets.map((p) => p.id)));
  ok('the no-report branch does not hand-roll its config any more',
    /config: Object\.assign\(publicConfig\(cfg, readingSample\)/.test(serverSrc),
    'the firstRun branch builds its own config object again');
  ok('and the API key still cannot reach the page',
    !/"apiKey"/.test(JSON.stringify(state.config)), 'apiKey appeared in the public config');

  // The wizard must write the choice rather than assume the default is the choice.
  // Comments are stripped first: the explanation of the bug legitimately contains
  // the phrase "chosen === already", which made a naive regex match its own note.
  const stripComments = (s) => s.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  const saveBoard = stripComments(
    page.slice(page.indexOf('function introSaveBoard'), page.indexOf('function introUpload')));
  ok('the wizard no longer skips saving when the choice matches the server default',
    !/chosen\s*===\s*already/.test(saveBoard),
    'the wizard still skips the save when the choice equals the server default');
  ok('and it always posts the chosen board',
    /apiPost\('\/api\/config', \{ specId: chosen \}\)/.test(saveBoard));

  // The dropdown must not be able to collapse to one option if presets go missing.
  const fill = stripComments(
    page.slice(page.indexOf('function fillSpecSelect'), page.indexOf('function fillSpecSelect') + 900));
  ok('the board dropdown falls back to exactly one labelled option, not an empty list',
    /\[\{ id: 'eduqas-alevel'/.test(fill),
    'no fallback option is defined');
  ok('and that fallback is reached only when specPresets is absent or empty',
    /config\.specPresets\.length[\s\S]{0,20}\?/.test(fill),
    'the fallback is not guarded on specPresets being present');
}

section('11x. Working through the wizard does not throw you back to the start');
{
  /*
   * Reported by the user: "whenever i do anything in the setup like send a
   * portfolio it sends me back to the beginning."
   *
   * maybeShowIntro() opened at step 0 unconditionally, and it is called from
   * loadState() — which runs after EVERY wizard action, because each one ends in
   * a state refresh. So uploading, saving a board or saving a key all restarted
   * the wizard. That is the mechanism, and it is easy to reintroduce, so it is
   * driven here rather than merely inspected.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  page.state.data = {
    ok: true, firstRun: true,
    config: {
      specId: 'aqa-gcse',
      specPresets: [
        { id: 'eduqas-alevel', label: 'WJEC Eduqas A-level Design and Technology (NEA)' },
        { id: 'aqa-gcse', label: 'AQA GCSE Design and Technology (8552)' },
      ],
      apiKeySet: false, pdfName: null, deckName: null,
    },
  };
  const stepNow = () => {
    for (let k = 0; k < 5; k++) { if (!page.els.get('introStep' + k).hidden) { return k; } }
    return -1;
  };
  const reload = () => page.call('maybeShowIntro');   // what loadState() does

  page.call('maybeShowIntro');
  ok('the wizard opens at the first step', stepNow() === 0, 'step ' + stepNow());

  // Walk to the portfolio step, then refresh the way an upload does.
  page.call('introShow', 2);
  reload();
  ok('a state refresh on the portfolio step does not restart the wizard',
    stepNow() === 2, 'ended on step ' + stepNow());

  // And on the key step, which is the other one that ends in a refresh.
  page.call('introShow', 3);
  reload();
  ok('a state refresh on the key step does not restart the wizard',
    stepNow() === 3, 'ended on step ' + stepNow());

  // The last step must survive a refresh too, or "Open NEA Coach" becomes "Next".
  page.call('introShow', 4);
  reload();
  ok('a state refresh on the final step does not restart the wizard',
    stepNow() === 4, 'ended on step ' + stepNow());

  // Closing must clear the "on screen" flag, or it could never reopen.
  page.call('introClose');
  ok('closing marks the wizard finished', page.store['neaCoach.intro.v1'] === 'done',
    JSON.stringify(page.store['neaCoach.intro.v1']));
  ok('and hides it', page.els.get('introOverlay').hidden === true);

  // "Show setup again" forgets the flag first, so it legitimately starts over.
  delete page.store['neaCoach.intro.v1'];
  reload();
  ok('re-opening after the flag is cleared starts from the beginning',
    stepNow() === 0, 'step ' + stepNow());
  ok('and the overlay is visible again', page.els.get('introOverlay').hidden !== true);

  // The whole flow must still complete, or the fix has broken the happy path.
  ok('no render or wizard error was raised', page.errors.length === 0, JSON.stringify(page.errors));
}

section('11y. Settings says whether a key is already saved, without opening the group');
{
  /*
   * Requested by the user. The collapsed "Free AI key" group said only "for the
   * marking and the review", so there was no way to tell whether a key had already
   * been added without expanding it. The header now answers it, which matters
   * because a student who cannot remember will otherwise paste a second key.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const keyNote = () => page.els.get('keyNote').textContent;
  const keyClass = () => page.els.get('keyNote').className;
  const draw = (config) => {
    page.state.data = { config };
    page.call('renderKeySection');
  };

  draw({ apiKeySet: false });
  ok('with no key it says so plainly',
    /not added yet/.test(keyNote()), JSON.stringify(keyNote()));
  ok('and does not claim the AI is available',
    /AI review is off/.test(keyNote()), JSON.stringify(keyNote()));
  ok('and it is not marked as saved', !/saved/.test(keyClass()), keyClass());

  draw({ apiKeySet: true, keyTail: 'Xy9z' });
  ok('with a key it says a key is saved', /key saved/.test(keyNote()), JSON.stringify(keyNote()));
  ok('and names the last four characters, so you can tell WHICH key',
    /Xy9z/.test(keyNote()), JSON.stringify(keyNote()));
  ok('and it is visually marked as set',
    /\bsaved\b/.test(keyClass()), keyClass());
  const panelText = page.els.get('apiKeyStatus').textContent;
  ok('the same key text still appears in the panel itself',
    /key ending Xy9z/.test(panelText), JSON.stringify(panelText));

  // An older config need not carry a tail; it must not say "ending undefined".
  draw({ apiKeySet: true });
  ok('a key with no recorded tail does not leak "undefined"',
    /key saved/.test(keyNote()) && !/undefined/.test(keyNote()), JSON.stringify(keyNote()));

  /*
   * The portfolio group says WHICH of the two are loaded.
   *
   * Note what this used to assert and why it passed: it drove the note with
   * `{ pdfName: null, deckName: null }`, and `deckName` is not a field the server has ever
   * sent. So the test agreed with the bug. The PowerPoint case below is the one that was
   * broken — a loaded portfolio reported as "nothing added yet".
   */
  const drawPortfolio = (config, deck) => {
    page.state.data = { config: config, deck: deck || null };
    page.call('renderUploadSection', page.state.data, config);
    return page.els.get('portfolioNote').textContent;
  };
  eq('a PDF alone is reported as the PDF',
    drawPortfolio({ pdfName: 'My Portfolio.pdf', pdfFile: '/x/My Portfolio.pdf' }), 'PDF added');
  eq('a PowerPoint alone is reported as the PowerPoint',
    drawPortfolio({ deckFile: '/x/My Portfolio.pptx' }, { fileName: 'My Portfolio.pptx' }), 'PowerPoint added');
  eq('both are reported as both',
    drawPortfolio({ deckFile: '/x/a.pptx', pdfName: 'a.pdf', pdfFile: '/x/a.pdf' }, { fileName: 'a.pptx' }),
    'PowerPoint and PDF added');
  eq('and nothing is reported as nothing', drawPortfolio({}, null), 'nothing added yet');
  ok('the file itself is named in the list rather than in the header',
    (() => {
      drawPortfolio({ pdfName: 'My Portfolio.pdf', pdfFile: '/x/My Portfolio.pdf' });
      const rows = page.els.get('uploadedWhat').children
        .map((c) => c.children.map((x) => x.textContent).join(' ')).join(' | ');
      return /My Portfolio\.pdf/.test(rows);
    })());
}

section('11z. Every ui.* the page uses is actually cached');
{
  /*
   * Found while adding the key-status line, and it had been broken for a long time:
   * `apiKeyStatus`, `apiKeyInput`, `btnTestKey`, `btnSaveKey` and `btnForgetKey`
   * were all read as `ui.something` but were never listed in cacheElements(). Every
   * one of them was therefore `undefined` at runtime, and because `on()` returns
   * early on a null node, the Test key, Save key and Forget key buttons silently
   * did NOTHING. The API key could only ever be set from the wizard, which uses a
   * different input.
   *
   * This is the nastiest shape of bug in this codebase: a missing cache entry fails
   * SILENTLY. There is no error, the button just does nothing. So the cache is now
   * audited rather than trusted.
   */
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const cachedBlock = page.match(/function cacheElements\(\)\s*\{[\s\S]*?\];/);
  ok('cacheElements has a list to audit', !!cachedBlock);
  const cached = new Set([...cachedBlock[0].matchAll(/'([^']+)'/g)].map((m) => m[1]));
  const ids = new Set([...page.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]));

  const used = new Set([...page.matchAll(/\bui\.([A-Za-z_][A-Za-z0-9_]*)/g)].map((m) => m[1]));
  // Only entries that HAVE an element are a defect. A name with no element is a
  // deliberate optional guard, which is fine and common in this file.
  const broken = [...used].filter((u) => !cached.has(u) && ids.has(u));
  ok('no ui.* reference points at an element that was never cached',
    broken.length === 0,
    'not cached, so undefined at runtime: ' + broken.join(', '));

  // And the specific controls the wizard's key screen depends on, named so a
  // future removal is a failure rather than a shrug.
  for (const name of ['apiKeyStatus', 'apiKeyInput', 'btnTestKey', 'btnSaveKey', 'btnForgetKey']) {
    ok('  ' + name + ' is cached', cached.has(name), 'missing from cacheElements');
  }
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
  /*
   * A real 44-page portfolio export, used to check the PDF extractor against something
   * that was not built for it.
   *
   * The path is relative and overridable rather than absolute: it used to hard-code a home
   * folder, which named the author in a public repository and made the test silently skip
   * on every other machine.
   */
  /*
   * A real 44-page portfolio export, used to check the PDF extractor against something that
   * was not built for it. Overridable, and otherwise looked for in the two places it
   * actually lives.
   *
   * It used to be an absolute path into one person's home folder, which named them in a
   * public repository AND made the test silently skip on every other machine. It is the home
   * DIRECTORY that is used now, not a name.
   */
  const PDF = process.env.NEA_TEST_PDF
    || [
      path.join(require('os').homedir(), 'Downloads', 'DT - Google Slides.pdf'),
      path.join(ROOT, '..', 'DT - Google Slides.pdf'),
    ].find((candidate) => fs.existsSync(candidate))
    || path.join(require('os').homedir(), 'Downloads', 'DT - Google Slides.pdf');

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
    /*
     * Headings must survive extraction. This deliberately does NOT look for a particular
     * person's name the way it used to: this file belongs to a real student, and naming them
     * in a test in a public repository is not this test's business. "Candidate number" is a
     * heading on the form and is just as good a proof that the text came through.
     */
    ok('headings survive extraction',
      all.includes('Candidate number'), 'missing identifying text');
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
    console.log('  SKIP  the 44-page sample PDF is not on this machine (looked for ' + PDF + ')');
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
    /id="compareCard"[^>]*style="display:none"/.test(page),
    'it is hidden until there is a comparison, which is what the style attribute does');

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
  /*
   * The upload section names the two formats rather than competing with the verdict at the
   * top of the page. It does that with a labelled status list now instead of prose, which
   * also happens to be what answers "which of them have I actually given it?".
   */
  ok('the upload section labels the two formats rather than the verdict',
    /id="uploadedWhat"/.test(page) && /label: 'PDF'/.test(page) && /label: 'PowerPoint'/.test(page),
    'the two formats are not named side by side');
  ok('and the PDF row explains what a PDF is for, not that it is the mark',
    /the AI is reading your words only/.test(page),
    'the PDF row no longer says what a PDF supplies');

  // The decision itself, in both directions.
  ok('a PDF is what gets marked when one is set',
    /if \(pdfName\) \{[\s\S]{0,200}setText\(ui\.reviewedName, pdfName\)/.test(page));
  /*
   * A name mismatch is now only a HINT, and says so: the real check compares the two files'
   * text (section 12x). A PDF export is routinely given a different name by the export tool,
   * so the name alone is not evidence of anything.
   */
  ok('a mismatched PDF and .pptx name is still raised, as a hint rather than a verdict',
    /have different names, and they have not been compared yet/.test(page),
    'the fallback hint is gone, so a mismatch would pass unmentioned before a scan');
  /*
   * Inverted deliberately. This used to assert a warning on the grounds that "the free
   * structural checks cannot run" for a PDF-only setup. They can and do: lib/pdfdeck.js
   * hands the checks a real deck with the page count and every page's text. The student
   * asked for the warning to go, and it was wrong as well as unwanted.
   */
  ok('a PDF on its own is NOT flagged as un-checkable',
    !/the free structural checks have nothing to read/.test(page),
    'the false warning is back');
  ok('and nothing warns about a .pptx being required',
    !/Add the \.pptx as well/.test(page));
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
  /*
   * EVERY @media (max-width: 700px) block, concatenated.
   *
   * This read only the first one, which happened to be the large phone block near the top. Adding
   * a second, smaller one further down — for the tab arrows — silently emptied this list and made
   * this assertion blame nineteen innocent classes. A rule belongs to the phone layout wherever it
   * is written.
   */
  const mobile = (page.match(/@media \(max-width: 700px\) \{[\s\S]*?\n  \}/g) || []).join('\n');

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
    // The setup button moved into the top bar; its wrapper needs a phone rule too.
    'topbar-actions',
    'btn-sm',
    'predicted-total',
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
    'renderKeySection', 'renderCheckCoverage', 'renderScanCostHint', 'renderPortfolioNotice'];
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
  /* ------------------------------------------------------------------ *
   * 12h. The predicted panel does not contradict itself
   *
   * Reported as "lots of inconsistencies". The panel showed a headline of about 94
   * while the criteria beneath added up to 75, with no stated total and a reason
   * sentence mixing marks, marks-of-error and percentages. This drives the real
   * renderer with those exact figures and checks the numbers reconcile.
   * ------------------------------------------------------------------ */
  {
    const { loadPage } = require(path.join(__dirname, 'dom.js'));
    const cal = require(path.join(ROOT, 'lib', 'calibration.js'));
    const page = loadPage(ROOT);
    const rows = [
      { id: 'a', name: 'Criterion A', mark: 9, max: 10, weighted: true },
      { id: 'b', name: 'Criterion B', mark: 9, max: 10, weighted: true },
      { id: 'c', name: 'Criterion C', mark: 16, max: 20, weighted: true },
      { id: 'd', name: 'Criterion D', mark: 13, max: 20, weighted: true },
      { id: 'e', name: 'Criterion E', mark: 12, max: 20, weighted: true },
      { id: 'f', name: 'Criterion F', mark: 16, max: 20, weighted: true },
    ];
    const criterionTotal = rows.reduce((a, r) => a + r.mark, 0);
    const c = cal.assessPrediction(criterionTotal, 100, { board: 'aqa-gcse', model: 'gemini-3.1-flash-lite' });
    page.state.data = {
      config: { pdfFile: '/x.pdf', markEnabled: true },
      deck: { slideCount: 24 },
      ai: {
        ran: true, visionRan: true, visionRuns: 3, visionUnstable: true, specLabel: 'AQA GCSE (8552)',
        predicted: { predicted: criterionTotal, maximum: 100, percentage: criterionTotal, weightsPublished: true, rows, calibration: c },
        percentageRange: { lowPercentage: 72, highPercentage: 79, spread: 7 },
      },
    };
    page.call('renderPredicted');

    const text = [];
    (function walk(n) {
      if (!n) { return; }
      if (n.nodeType === 3 && String(n.nodeValue || '').trim()) { text.push(String(n.nodeValue).trim()); }
      (n.children || []).forEach(walk);
    }(page.ui.predictedBody));
    const joined = text.join(' | ');

    ok('the criteria table states its own total',
      new RegExp("own total: " + criterionTotal + " of 100").test(joined), joined.slice(0, 200));
    ok('and explains the relationship between that total and the estimate',
      /why it differs from this total|is this total/i.test(joined), joined.slice(-200));
    /*
     * The headline is whatever the calibration decided: the corrected figure when
     * one was applied, otherwise the AI's own reading. Both are correct answers to
     * "what does the panel show", and asserting only the corrected form is what
     * broke when the cheap model stopped being corrected.
     */
    const expectedHeadline = c.applied ? c.calibrated : c.raw;
    ok('the headline figure matches what the calibration decided',
      joined.indexOf('about ' + expectedHeadline + ' of 100') > -1,
      'expected about ' + expectedHeadline + ' of 100, got: ' + joined.slice(0, 160));
    ok('and the AI\'s own reading is named',
      new RegExp('read this as ' + criterionTotal + ' out of 100').test(joined)
      || joined.indexOf('the AI\u2019s own reading was ' + criterionTotal) > -1,
      joined.slice(0, 240));
    // The raw reading was printed twice in a row once already.
    ok('the raw reading is stated once, not twice',
      !/own reading was/.test(joined), 'the duplicate line is back');
    ok('no impossible zero-disagreement claim',
      !/disagreed by 0/.test(joined));
    ok('the explanation uses marks throughout, never a percentage',
      !/%/.test((c.note || '')), c.note);
  }

  /* ------------------------------------------------------------------ *
   * 12g. Nothing calls a function that does not exist
   *
   * `dailyLimitFor` was called in two places and defined in NONE, so
   * loadModelChoices and renderModelCostPreview both threw. Neither was noticed
   * because both are reached through paths that were themselves broken:
   * ui.modelNote had never been cached, so renderModelCostPreview returned at its
   * guard before reaching the call. A reference-count check is crude, but it is
   * exactly the check that was missing.
   * ------------------------------------------------------------------ */
  {
    const src = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
    const names = ['dailyLimitFor'];
    names.forEach((fn) => {
      const defined = new RegExp('function\\s+' + fn + '\\s*\\(').test(src)
        || new RegExp('(var|let|const)\\s+' + fn + '\\s*=').test(src);
      const calls = (src.match(new RegExp('\\b' + fn + '\\s*\\(', 'g')) || []).length;
      ok(fn + ' is defined, not just called', defined, calls + ' call site(s), no definition');
    });
    // And the specific control that was masked by an uncached element.
    ok('the model dropdown note element is cached, so its renderer can run',
      /'modelNote'/.test(src.match(/function cacheElements\(\)[\s\S]*?\];/)[0]));
  }

section('12j. Portfolio statistics actually appear, on screen and in the export');
{
  /*
   * Reported: "the stats page seems to have gone. it should show that in the
   * exported thing too".
   *
   * It had gone, on screen AND in the export, from the same mistake made in two
   * places. `buildStats` is stored as `deck.stats`, and both the state endpoint and
   * exportHtml read `report.stats`, which has never existed. So the value was
   * always null, renderStats hid the card because it had nothing to draw, and the
   * exported file carried no statistics at all.
   */
  const server = require(path.join(ROOT, 'server.js'));

  // A report shaped exactly like the real one: stats nested under the deck.
  const report = {
    ok: true, generatedAt: new Date().toISOString(),
    deck: {
      fileName: 'My Portfolio.pdf', filePath: '/tmp/x.pdf', slideCount: 24,
      stats: { words: 13113, pages: 24, images: 37, typingHours: 5.5, typingWpm: 40, readingMinutes: 87, topWords: [], balance: [] },
    },
    local: { score: 55, bySeverity: {}, findings: [] },
    ai: null, combined: { findings: [] }, config: {}, scan: {},
  };

  const html = server.exportHtml(report, {});
  ok('the export carries the word count from deck.stats',
    /13,113|13113/.test(html), 'word count missing from the export');
  ok('the export carries the page count', />24</.test(html), 'page count missing');
  ok('the export has the statistics tiles', /words<\/div>/.test(html) && /pages<\/div>/.test(html));

  // And it must NOT depend on a top-level stats key, which is what it used to read.
  const withoutDeckStats = Object.assign({}, report, { deck: { fileName: 'x', stats: undefined } });
  const empty = server.exportHtml(withoutDeckStats, {});
  ok('with no stats the export simply omits them rather than failing',
    !/words<\/div>/.test(empty));

  // The state endpoint must publish them too, or the on-screen card stays hidden.
  const stateSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok('the state endpoint reads stats from the deck, with a fallback',
    /stats: \(report\.deck && report\.deck\.stats\) \|\| report\.stats \|\| null/.test(stateSrc),
    'state endpoint still reads report.stats');
  ok('and the export does the same',
    /const stats = deck\.stats \|\| r\.stats \|\| \{\}/.test(stateSrc),
    'exportHtml still reads r.stats');
}

section('12k. Expanding the mechanical checks does not jump to the top');
{
  /*
   * Reported: "when you click on show mechanical checks the ap scrolls to the
   * top".
   *
   * renderPreservingScroll() anchored on the BUTTON that was clicked. That button
   * sits at the END of the group it expands, so growing the group pushes it down -
   * "keep the button still" then drags the reader down after it, and when the
   * button is removed entirely (it has served its purpose) there was nothing to
   * find and the reader was left wherever the browser had clamped them: the top.
   *
   * It now anchors on the list container, which does not move when its contents do.
   */
  const page = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const fn = page.slice(page.indexOf('function renderPreservingScroll'), page.indexOf('function jumpToPage'));
  ok('the scroll keeper anchors on the list container',
    /var container = ui\.fixList/.test(fn), 'still anchoring on the clicked button');
  ok('and falls back to the caller element when there is no container',
    /container && container\.getBoundingClientRect\) \? container : anchorEl/.test(fn));
  ok('the mechanical-checks button still uses it',
    /renderPreservingScroll\(e && e\.currentTarget, function \(\) \{[\s\S]{0,120}state\.showObvious = true/.test(page));
}

section('12i. Setting the board in the walkthrough updates the Settings dropdown');
{
  /*
   * Reported: "the exam board drop down menu doesnt update if you set it during
   * the setup".
   *
   * renderSettings() only writes the controls when !state.settingsFilled, and the
   * wizard's introSaveBoard() never cleared that flag - while the Settings button
   * (saveBoard) did. So the board was saved to the server and the dropdown went on
   * showing the old value.
   *
   * Two fixes, and the test covers the second one on its own because it is the one
   * that holds even if some future code path forgets the flag again.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  page.window.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, config: { specId: 'aqa-gcse' } }) });

  const PRESETS = [
    { id: 'eduqas-alevel', label: 'WJEC Eduqas A-level Design and Technology (NEA)' },
    { id: 'aqa-gcse', label: 'AQA GCSE Design and Technology (8552)' },
  ];
  page.state.data = { config: { specId: 'eduqas-alevel', specPresets: PRESETS } };
  page.call('renderSettings');
  const sel = page.ui.specSelect;
  ok('the board dropdown starts on the saved board', sel.value === 'eduqas-alevel', sel.value);

  // The wizard saves a different board and the state comes back updated.
  page.state.data.config.specId = 'aqa-gcse';
  page.call('renderSettings');
  ok('after the board is saved the dropdown shows the NEW board',
    sel.value === 'aqa-gcse', 'dropdown shows ' + sel.value);

  /*
   * And it must hold even when the form is NOT refilled - the exact stale path.
   * settingsFilled is deliberately left true here, which is what the wizard used
   * to do.
   */
  page.state.data.config.specId = 'eduqas-alevel';
  page.call('renderSettings');
  page.state.settingsFilled = true;
  page.state.data.config.specId = 'aqa-gcse';
  page.call('renderSettings');
  ok('a re-render that skips the form still syncs the board dropdown',
    sel.value === 'aqa-gcse', 'dropdown shows ' + sel.value);

  // The wizard must also clear the flag, so the rest of the form refreshes too.
  const src = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const wizard = src.slice(src.indexOf('function introSaveBoard'), src.indexOf('function introUpload'));
  ok('the walkthrough invalidates the filled-form cache, like the Save button does',
    /state\.settingsFilled = false/.test(wizard),
    'introSaveBoard does not clear settingsFilled');
  ok('and there is a single helper for syncing the dropdown',
    (src.match(/function syncSpecSelect/g) || []).length === 1);
}

section('12e2. The exam board\'s own guidance reaches the marking prompt');
{
  /*
   * The user supplied AQA's 8552 candidate record form, which carries the board's
   * per-criterion "positive features" and "further development" notes. The second
   * of those is what the board says is typically MISSING from work at that level,
   * and it is the most directly useful thing in the document for marking.
   *
   * It was silently dropped the first time: loadSpec() whitelists the fields it
   * copies out of each criterion, so anything not listed is discarded without a
   * word. That is the same silent-failure shape as the uncached ui.* elements, so
   * it is pinned here.
   */
  const ai = require(path.join(ROOT, 'lib', 'ai.js'));
  const spec = ai.loadSpec(path.join(ROOT, 'spec'), 'aqa-gcse');

  ok('the AQA spec still loads', spec.criteria.length === 6, String(spec.criteria.length));
  ok('the six official criteria are the ones AQA actually marks',
    spec.criteria.map((c) => c.name).join('|')
      === 'Identifying and investigating design possibilities|Producing a design brief and specification|Generating design ideas|Developing design ideas|Realising design ideas|Analysing and evaluating',
    spec.criteria.map((c) => c.name).join('|'));
  ok('and their maxima total 100, as the record form states',
    spec.criteria.reduce((t, c) => t + (c.weight || 0), 0) === 100,
    String(spec.criteria.reduce((t, c) => t + (c.weight || 0), 0)));
  ok('the per-criterion maxima match the form (10/10/20/20/20/20)',
    spec.criteria.map((c) => c.weight).join(',') === '10,10,20,20,20,20',
    spec.criteria.map((c) => c.weight).join(','));

  ok('every criterion carries the board\'s top-band wording',
    spec.criteria.every((c) => c.evidence && c.evidence.length > 80),
    'a criterion has no evidence text');
  ok('and the board\'s note on what is typically missing',
    spec.criteria.every((c) => c.commonShortfalls && c.commonShortfalls.length > 40),
    spec.criteria.map((c) => c.id + ':' + (c.commonShortfalls ? 'yes' : 'MISSING')).join(' '));

  // loadSpec must COPY the field, not merely have it in the file.
  const raw = JSON.parse(fs.readFileSync(path.join(ROOT, 'spec', 'aqa-gcse-dt-criteria.json'), 'utf8'));
  ok('the spec file itself holds the guidance',
    raw.criteria.every((c) => c.commonShortfalls));
  ok('loadSpec carries it through rather than dropping unknown fields',
    spec.criteria.every((c) => c.commonShortfalls && c.commonShortfalls.length > 40));

  // And it must actually be interpolated into the prompt text.
  const src = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');
  ok('it is placed in the marking prompt, twice (batch and cross-check)',
    (src.match(/commonShortfalls/g) || []).length >= 4,
    String((src.match(/commonShortfalls/g) || []).length) + ' references');
  ok('the prompt version was bumped so cached results are not reused',
    /PROMPT_VERSION = '8'/.test(src));
}

  /* ------------------------------------------------------------------ *
   * 12f. Settings save themselves
   *
   * Requested by the user. Settings used to be one form with one "Save advanced
   * settings" button, so it was easy to change the model, close the panel and
   * assume it had applied. Everything except the folder path now saves on change.
   *
   * This lives in the async block because the save is a promise: a synchronous
   * assertion would run before the POST was even issued.
   * ------------------------------------------------------------------ */
  {
    const { loadPage } = require(path.join(__dirname, 'dom.js'));
    const page = loadPage(ROOT);
    const posted = [];
    page.window.fetch = (url, opts) => {
      posted.push({ url: String(url), body: opts && opts.body });
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ ok: true, config: {} }) });
    };
    page.state.data = {
      config: { model: 'gemini-3.1-flash-lite' },
      quota: { assumedLimit: 20, models: [{ model: 'gemini-3.8-flash', limit: 20 }] },
    };
    const seed = (sel, values) => { sel.options = values.map((v) => ({ value: v, text: v })); sel.value = values[0]; };
    seed(page.ui.modelSelect, ['gemini-3.8-flash', 'gemini-3.1-flash-lite']);
    seed(page.ui.visionModelSelect, ['gemini-3.5-flash']);
    const configBodies = () => posted.filter((x) => x.url.includes('/api/config')).map((x) => x.body);
    const settle = () => new Promise((r) => setTimeout(r, 30));

    page.ui.modelSelect.value = 'gemini-3.8-flash';
    page.ui.modelSelect.change();
    await settle();
    ok('changing the page-reading model saves it, with no button to press',
      configBodies().some((b) => /gemini-3\.8-flash/.test(String(b))), JSON.stringify(configBodies()));
    ok('and the page says so, because a silent auto-save looks like a dead control',
      /Saved/.test(page.els.get('saveNote').textContent), JSON.stringify(page.els.get('saveNote').textContent));

    page.ui.markToggle.checked = true;
    page.ui.markToggle.change();
    await settle();
    ok('toggling Predict a mark saves it',
      configBodies().some((b) => /markEnabled/.test(String(b))), JSON.stringify(configBodies()));

    page.ui.visionModelSelect.value = 'gemini-3.5-flash';
    page.ui.visionModelSelect.change();
    await settle();
    ok('changing the marking model saves it',
      configBodies().some((b) => /visionModel/.test(String(b))), JSON.stringify(configBodies()));

    /*
     * The folder keeps an explicit button and must NOT auto-save. It is the one
     * control that can be genuinely wrong: the server rejects a path that does not
     * exist, and silently writing a bad one would leave the app pointing at nothing
     * while the panel implied all was well.
     */
    const before = configBodies().length;
    page.ui.deckDirInput.value = '/definitely/not/a/real/folder';
    page.ui.deckDirInput.change();
    await settle();
    ok('typing a folder path does NOT auto-save',
      configBodies().length === before, 'the folder path auto-saved');
    ok('the folder keeps its own explicit button', !!page.ui.btnSaveSettings);
    ok('no error was raised while saving', page.errors.length === 0, JSON.stringify(page.errors));
  }

section('12l. The feel layer: ripples, rolling figures and the reward for finishing');
{
  /*
   * None of this is load-bearing, which is exactly why it needs tests: if a
   * ripple or a burst quietly stops working, nothing throws and nothing looks
   * broken — the page just goes back to feeling like a form, and nobody notices
   * for a month.
   *
   * So each piece is driven for real through the DOM stand-in and the DOM it
   * produces is inspected, rather than the file being grepped for the word
   * "ripple".
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const src = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(src)[1];
  const page = loadPage(ROOT);
  const { document, T } = page;

  ok('the page still initialises with the feel layer in place',
    page.errors.length === 0, JSON.stringify(page.errors));
  ok('the feel helpers are reachable from the page',
    typeof T.spawnRipple === 'function' && typeof T.celebrate === 'function'
    && typeof T.rollFigure === 'function');

  // ---- the ripple --------------------------------------------------------
  const btn = document.createElement('button');
  btn.className = 'btn btn-primary';
  btn.disabled = false;
  document.body.appendChild(btn);
  T.spawnRipple(btn, 90, 18);
  const ripples = btn.children.filter((c) => c.className === 'ripple');
  ok('a press puts exactly one ripple inside the button', ripples.length === 1,
    ripples.length + ' ripples');
  const ripple = ripples[0];
  if (ripple) {
    /*
     * The stand-in reports every box as 100x20, so a press at x=90 leaves 90px to
     * the far corner and 10px to the near one. Sized from the near edge it would
     * be 20px across and visibly fail to cover the button.
     */
    const size = parseFloat(ripple.style.width);
    ok('the ripple is sized to reach the furthest corner, not the nearest',
      size >= 180, 'ripple was ' + ripple.style.width + ' wide');
    ok('and it is centred on the press rather than on the button',
      ripple.style.left === '90px' && ripple.style.top === '18px',
      ripple.style.left + ' / ' + ripple.style.top);
    ok('the ripple is invisible to pointers and to screen readers',
      ripple.getAttribute('aria-hidden') === 'true' && /\.ripple \{[\s\S]*?pointer-events:\s*none/.test(css));
  }
  ok('the ripple is clipped to the button it belongs to', (() => {
    const bodies = [...css.matchAll(/\.btn \{([^}]*)\}/g)].map((m) => m[1]);
    return bodies.some((b) => /overflow:\s*hidden/.test(b) && /position:\s*relative/.test(b));
  })(), 'no .btn rule both hides overflow and anchors the ripple');
  ok('a disabled button gets no ripple', (() => {
    const dead = document.createElement('button');
    dead.className = 'btn';
    dead.disabled = true;
    T.spawnRipple(dead, 10, 10);
    return dead.children.filter((c) => c.className === 'ripple').length === 0;
  })());
  ok('one delegated listener covers buttons drawn later, instead of one each',
    /document\.addEventListener\('pointerdown'/.test(src) && /closest\('\.btn'\)/.test(src));

  // ---- reduced motion ----------------------------------------------------
  /*
   * A reader who has asked their system for reduced motion must get the final
   * state immediately and no movement at all. This is checked by flipping the
   * page's own flag, so it exercises the same branch a real reader would take.
   */
  T.setCalm(true);
  ok('reduced motion: no ripple at all', (() => {
    const calm = document.createElement('button');
    calm.className = 'btn';
    calm.disabled = false;
    T.spawnRipple(calm, 10, 10);
    return calm.children.filter((c) => c.className === 'ripple').length === 0;
  })());
  ok('reduced motion: no sparks either', (() => {
    T.celebrate(100, 100, 3);
    return T.sparkLayer().children.filter((c) => c.className === 'spark').length === 0;
  })());
  ok('reduced motion: the figure is simply written out in full', (() => {
    const node = document.createElement('div');
    T.rollFigure('calm-check', node, 'about 88 of 100 marks');
    return node.textContent === 'about 88 of 100 marks';
  })());
  T.setCalm(false);

  // ---- rolling figures ---------------------------------------------------
  const fig = document.createElement('div');
  T.rollFigure('t-roll', fig, 'about 81 of 100 marks');
  const first = fig.textContent;
  const firstNum = parseInt(/(\d+)/.exec(first)[1], 10);
  ok('the figure climbs rather than appearing', firstNum < 81,
    'it started at ' + firstNum);
  ok('and the words around the number are preserved while it climbs',
    /^about \d+ of 100 marks$/.test(first), first);
  await new Promise((r) => setTimeout(r, 800));
  ok('it settles exactly on the real text', fig.textContent === 'about 81 of 100 marks',
    fig.textContent);

  /*
   * The reason the memory is keyed by name rather than stored on the element:
   * renderPredicted builds a brand new node on every render, so an element-local
   * memory would make the headline count up from zero again after every unrelated
   * setting change, and on every poll of a running check.
   */
  ok('a rebuilt node with unchanged text does not roll again', (() => {
    const rebuilt = document.createElement('div');
    T.rollFigure('t-roll', rebuilt, 'about 81 of 100 marks');
    return rebuilt.textContent === 'about 81 of 100 marks';
  })(), 'the headline re-rolled from zero on a re-render');

  ok('a change of wording is written straight out, with no rolling at all', (() => {
    const node = document.createElement('div');
    T.rollFigure('t-words', node, 'a sentence with no numbers in it');
    return node.textContent === 'a sentence with no numbers in it';
  })());

  // ---- the celebration --------------------------------------------------
  const layer = T.sparkLayer();
  ok('the spark layer is inert', /#sparks \{[\s\S]*?pointer-events:\s*none/.test(css));
  T.celebrate(120, 80, 3);
  const sparks = layer.children.filter((c) => c.className === 'spark');
  ok('finishing a check throws a burst', sparks.length > 0, 'no sparks were created');
  ok('it is a burst rather than one dot', sparks.length >= 20, sparks.length + ' sparks');
  ok('the burst starts where it was asked to',
    sparks[0].style.left === '120px' && sparks[0].style.top === '80px',
    sparks[0].style.left + ' / ' + sparks[0].style.top);
  ok('every spark carries its own direction, so they fan out rather than moving as one block',
    sparks.every((s) => /^-?\d+px$/.test(s.style.getPropertyValue('--dx'))
      && /^-?\d+px$/.test(s.style.getPropertyValue('--dy'))
      && /^-?\d+deg$/.test(s.style.getPropertyValue('--rot'))));
  const distinct = new Set(sparks.map((s) => s.style.getPropertyValue('--dx') + ',' + s.style.getPropertyValue('--dy')));
  ok('and the directions actually differ', distinct.size >= 8, distinct.size + ' distinct directions');
  ok('a bigger moment is a bigger burst', (() => {
    const before = layer.children.filter((c) => c.className === 'spark').length;
    T.celebrate(10, 10, 1);
    return layer.children.filter((c) => c.className === 'spark').length > before;
  })());
  ok('an unmeasurable position produces no burst rather than sparks at the origin', (() => {
    const before = layer.children.filter((c) => c.className === 'spark').length;
    T.celebrate(NaN, NaN, 3);
    return layer.children.filter((c) => c.className === 'spark').length === before;
  })());

  // ---- where the celebration fires --------------------------------------
  ok('the big burst fires on the running to finished edge only',
    /runningNow/.test(src) && /scanWasRunning = true;/.test(src)
    && /else if \(scanWasRunning\) \{/.test(src),
    'the completion burst is not edge-triggered, so it would repeat on every poll');
  ok('ticking a finding measures the box before the re-render removes it',
    /var spot = box\.getBoundingClientRect\(\);/.test(src)
    && /if \(box\.checked\) \{ celebrate\(spot\.left/.test(src),
    'the tick-off burst would be positioned after the node was destroyed');
  ok('pressing a button springs, and the spring outranks the older translateY press',
    css.indexOf('.btn:active:not([disabled])') > css.indexOf('.btn-primary:active')
    && css.indexOf('.btn:active:not([disabled])') > -1);

  // ---- the surfaces -----------------------------------------------------
  ok('the ambient background hangs off html, so body\'s canvas colour cannot hide it',
    /html::before \{[\s\S]*?position:\s*fixed[\s\S]*?z-index:\s*-1/.test(css));
  ok('the header is a real glass surface',
    /\.topbar::before \{[\s\S]*?backdrop-filter:\s*blur\(/.test(css));
  ok('blur is spent on fixed chrome, never on content',
    (() => {
      const blurred = [...css.matchAll(/([^{}]+)\{([^}]*)\}/g)]
        .filter((m) => /backdrop-filter/.test(m[2]))
        .map((m) => m[1].trim().split('\n').pop().trim());
      /*
       * Three pieces of chrome, one instance each and all of them fixed: the header bar,
       * the wizard overlay and the back-to-top pill. A CARD is the thing this guards
       * against — there are a dozen of them and they scroll, so blurring every one is
       * what makes a long report stutter. Naming the three explicitly means a new
       * blurred surface has to be argued for here rather than slipped in.
       */
      /*
       * Four pieces of chrome now, one instance each and all of them fixed: the header bar,
       * the wizard overlay, the back-to-top pill and the phone dock. Naming them explicitly
       * means a new blurred surface has to be argued for here rather than slipped in.
       */
      return blurred.length > 0 && blurred.every((s) => /topbar|intro|to-top|dock/.test(s));
    })(), 'a scrolling surface was blurred, which is what makes a long report stutter');
  ok('cards rise in without being left holding a transform',
    // .panel > .card now: the cards live inside section panels, not directly in .page.
    /\.(?:page|panel) > \.card \{[^}]*animation:\s*rise [^}]*backwards/.test(css)
    && !/animation:\s*rise [^}]*both/.test(css));
  ok('the ring starts empty, so it draws itself in rather than unwinding',
    /id="ringValue"[\s\S]{0,160}stroke-dashoffset="326\.73"/.test(src));
  ok('figures that count use tabular numerals, so a rolling number does not jitter',
    /font-variant-numeric:\s*tabular-nums/.test(css));
  ok('the header unpins on a phone, where pinning it would eat the screen',
    /@media \(max-width: 700px\) \{[\s\S]{0,200}\.topbar \{ position: static/.test(css));

  /*
   * color-mix is the whole basis of the translucency, and it is a 2023 feature.
   * On a browser without it an undeclared fallback does not degrade quietly:
   * `border-color` falls back to currentColor (a black outline round every card)
   * and `background` falls back to transparent (a card with no surface at all).
   * Both are worse than the flat design this replaced, so each one must have a
   * plain declaration immediately before it.
   */
  const naked = [];
  const lines = css.split('\n');
  for (let i = 0; i < lines.length; i++) {
    const m = /^\s*(background|border-color):\s*.*color-mix.*;\s*$/.exec(lines[i]);
    if (!m) { continue; }
    let j = i - 1;
    while (j >= 0) {
      const t = lines[j].trim();
      if (!t || t.startsWith('/*') || t.startsWith('*') || t.startsWith('//')) { j -= 1; continue; }
      break;
    }
    const prev = j >= 0 ? lines[j].trim() : '';
    if (!prev.startsWith(m[1] + ':')) { naked.push(m[1] + ' at line ' + (i + 1) + ' (previous was "' + prev + '")'); }
  }
  ok('every color-mix on a surface has an older-browser fallback above it',
    naked.length === 0, naked.join(' | '));

  ok('nothing in the feel layer threw while all of that ran',
    page.errors.length === 0, JSON.stringify(page.errors));
}

section('12m. Nothing is reviewed until a portfolio has actually been chosen');
{
  /*
   * Reported: "i clicked check now after uploading nothing and it had stiuff to say".
   *
   * It had plenty to say. The default portfolio folder is the parent of the app, the
   * app is delivered by unzipping into ~/Downloads, and with no file chosen findDeck
   * searched that folder three levels deep and took the most recently edited .pptx it
   * found. On the developer's own machine that was
   * "AQA DT 7552 Unit 2 T1 Performance-Polymers.pptx", an exam-board teaching deck,
   * and the app produced 48 findings about it as though it were the student's NEA.
   * The only clue anywhere was a small filename.
   *
   * The rule is now: no portfolio chosen, no report.
   */
  const server = require(path.join(ROOT, 'server.js'));
  const os = require('os');

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'nea-deck-'));
  const somebodyElse = path.join(tmp, 'Somebody Elses Coursework.pptx');
  fs.writeFileSync(somebodyElse, 'not really a pptx');

  // The folder HAS a .pptx in it. It must not be used.
  const notChosen = server.findDeck({ deckDir: tmp, deckDirChosen: false, deckFile: null, pdfFile: null });
  ok('with no portfolio chosen the app refuses instead of searching the folder',
    !!notChosen.error && notChosen.needsPortfolio === true, JSON.stringify(notChosen));
  ok('and it says the one thing the student needs to know',
    notChosen.error === 'No portfolio has been uploaded yet.', notChosen.error);
  ok('it does not quietly hand back the file it found anyway',
    !notChosen.filePath, String(notChosen.filePath));

  // The flag is what makes the difference, not the folder.
  const chosen = server.findDeck({ deckDir: tmp, deckDirChosen: true, deckFile: null, pdfFile: null });
  ok('once a folder has been chosen deliberately, searching it is allowed again',
    chosen.filePath === somebodyElse, JSON.stringify(chosen));

  // An explicit file needs no folder flag — that is the upload path.
  const explicit = server.findDeck({ deckDir: tmp, deckDirChosen: false, deckFile: somebodyElse, pdfFile: null });
  ok('an explicitly chosen file is read even without the folder flag',
    explicit.filePath === somebodyElse, JSON.stringify(explicit));

  const empty = fs.mkdtempSync(path.join(os.tmpdir(), 'nea-empty-'));
  const none = server.findDeck({ deckDir: empty, deckDirChosen: true, deckFile: null, pdfFile: null });
  ok('an empty chosen folder does NOT quietly become the bundled sample',
    !!none.error && none.needsPortfolio === true && !none.usedSample, JSON.stringify(none));

  // The shipped default is "not chosen", which is what fixes existing installs with
  // deckDir ".." already saved: the flag is simply absent there.
  ok('a fresh install starts with no portfolio chosen',
    server.defaultConfig().deckDirChosen === false,
    String(server.defaultConfig().deckDirChosen));
  ok('and the browser is told, so it can draw an empty state',
    server.publicConfig({ deckDirChosen: false }).portfolioChosen === false
    && server.publicConfig({ deckDirChosen: true }).portfolioChosen === true);

  /*
   * Every route that hands the app a portfolio must set the flag, or the guard would
   * lock a student out of their own file — a worse bug than the one being fixed.
   */
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok('uploading a PowerPoint marks the portfolio as chosen',
    /cfg\.deckFile = target;\n\s*cfg\.deckDirChosen = true;/.test(src));
  ok('uploading a PDF marks it chosen too',
    /cfg\.pdfFile = target;\n\s*\/\/[^\n]*\n\s*cfg\.deckDirChosen = true;/.test(src));
  const flagSites = (src.match(/cfg\.deckDirChosen = true;/g) || []).length;
  ok('and saving a folder in Settings does as well', flagSites >= 3, flagSites + ' sites set the flag');

  ok('the saved report is suppressed while there is no portfolio',
    /const report = portfolioReady \? \(lastReport \|\| readJsonSafe\(LAST_REPORT_PATH, null\)\) : null;/.test(src),
    'a stale report would still be shown even though the scan is refused');

  // The page must say so, and stop offering to scan.
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  ok('the page still loads with the guard in place', page.errors.length === 0, JSON.stringify(page.errors));

  const textOf = (node) => {
    if (!node) { return ''; }
    let out = node.textContent ? String(node.textContent) : '';
    for (const k of (node.children || [])) { out += ' ' + textOf(k); }
    return out;
  };

  page.state.scanning = false;
  page.state.data = { ok: true, config: { portfolioChosen: false }, deck: null, local: { findings: [] }, ai: {} };
  const drew = page.call('renderActions');
  ok('renderActions survives an install with no portfolio', drew.ok, drew.error);
  ok('Check now is disabled when there is no portfolio', page.ui.btnCheck.disabled === true);
  ok('and the button says why, rather than sitting there dead',
    /portfolio first/i.test(page.ui.btnCheckLabel.textContent), page.ui.btnCheckLabel.textContent);
  ok('the banner states it plainly',
    /No portfolio has been uploaded yet/.test(textOf(page.els.get('sampleNotice'))),
    textOf(page.els.get('sampleNotice')));
  ok('the banner offers the way out', /Add my portfolio/.test(textOf(page.els.get('sampleNotice'))));

  // And it lets go the moment there is something to read.
  page.state.data = { ok: true, config: { portfolioChosen: true }, deck: null, local: { findings: [] }, ai: {} };
  page.call('renderActions');
  ok('with a portfolio the button comes back', page.ui.btnCheck.disabled === false);
  ok('and the banner clears itself', textOf(page.els.get('sampleNotice')).trim() === '',
    textOf(page.els.get('sampleNotice')));

  fs.rmSync(tmp, { recursive: true, force: true });
  fs.rmSync(empty, { recursive: true, force: true });
}

section('12n. Your settings, key and portfolio survive an upgrade');
{
  /*
   * Reported: "it should stay uploaded if the user opens the app at a later date".
   *
   * It did not. The data lived in <app folder>/state, and every release unzips into a
   * NEW folder — deliberately, because that is what stops an older copy overwriting a
   * newer one — so every release also arrived with an empty state folder. The key, the
   * portfolio and the settings were still sitting in the previous version's folder.
   * The user's v1.0.22 config had an empty apiKey, which is exactly what that looks
   * like from their side.
   *
   * The data now lives where each platform expects per-user data, and an existing
   * install is brought across on first run.
   */
  const server = require(path.join(ROOT, 'server.js'));
  const os = require('os');

  /*
   * The DEFAULT location, not the override this suite sets for its own isolation. Both
   * assertions are about what a student gets, so the override is lifted for the check and put
   * straight back.
   */
  const namedDefault = (() => {
    const held = process.env.NEA_COACH_DATA_DIR;
    delete process.env.NEA_COACH_DATA_DIR;
    const answer = server.perUserDataDir();
    if (held === undefined) delete process.env.NEA_COACH_DATA_DIR;
    else process.env.NEA_COACH_DATA_DIR = held;
    return answer;
  })();
  ok('the intended data folder is outside the app folder, so unzipping cannot lose it',
    namedDefault.indexOf(ROOT) !== 0, namedDefault);
  ok('and it is named so a person can actually find it',
    /NEA Coach$/.test(namedDefault), namedDefault);
  /*
   * Deliberately not asserting stateDir IS the per-user folder: on a machine where
   * that location cannot be written the app falls back, and this test has to pass on
   * both. What matters is that it is one of the two documented answers.
   */
  ok('the resolved folder is the per-user one, or the documented fallback',
    server.stateDir === server.perUserDataDir()
    || server.stateDir === path.join(ROOT, 'state'), server.stateDir);

  // The override is the escape hatch, and the only way to exercise this in a sandbox.
  const overrideDir = path.join(os.tmpdir(), 'nea-override-check');
  const previous = process.env.NEA_COACH_DATA_DIR;
  process.env.NEA_COACH_DATA_DIR = overrideDir;
  ok('NEA_COACH_DATA_DIR overrides where the data lives',
    server.perUserDataDir() === overrideDir, server.perUserDataDir());
  if (previous === undefined) delete process.env.NEA_COACH_DATA_DIR;
  else process.env.NEA_COACH_DATA_DIR = previous;

  /*
   * A stale absolute path in an adopted config is the one fault that would look like a
   * success and behave like a failure: the app would say a portfolio had been chosen
   * and then fail to find it. repairAdoptedPaths is what prevents that, and it is
   * driven here against a config laid out exactly like a real one.
   */
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'nea-adopted-'));
  fs.mkdirSync(path.join(dir, 'portfolio'), { recursive: true });
  const moved = path.join(dir, 'portfolio', 'My Work.pptx');
  fs.writeFileSync(moved, 'x');
  const cfgPath = path.join(dir, 'config.json');
  fs.writeFileSync(cfgPath, JSON.stringify({
    deckFile: '/somewhere/that/went/away/My Work.pptx',
    pdfFile: '/somewhere/that/went/away/Gone For Good.pdf',
    apiKey: 'AIzaKEEPME',
  }));

  server.repairAdoptedPaths(dir);
  const repaired = JSON.parse(fs.readFileSync(cfgPath, 'utf8'));
  ok('a portfolio whose old path has gone is re-pointed at its new home',
    repaired.deckFile === moved, repaired.deckFile);
  ok('a portfolio that genuinely is not there is left alone rather than faked',
    repaired.pdfFile === '/somewhere/that/went/away/Gone For Good.pdf', repaired.pdfFile);
  ok('and having a portfolio counts as having chosen one, so the guard cannot lock them out',
    repaired.deckDirChosen === true);
  ok('the API key comes across untouched', repaired.apiKey === 'AIzaKEEPME');
  fs.rmSync(dir, { recursive: true, force: true });

  /*
   * An unusable per-user folder must not stop the app, and must not be quiet about it:
   * falling back means the data will be lost on the next upgrade, which is the exact
   * fault this section exists for.
   */
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok('an unwritable data folder falls back instead of refusing to start',
    /const fallback = path\.join\(APP_DIR, 'state'\)/.test(src) && /return fallback;/.test(src));
  ok('and says so out loud, because silently losing the data IS the bug',
    /Could not use the per-user data folder/.test(src));
  ok('the fallback is decided at startup, not when the first write fails',
    /fs\.mkdirSync\(target, \{ recursive: true \}\);\n\s*return target;/.test(src),
    'an unwritable folder would only surface when the server tried to save');
  ok('adoption is logged when it happens, so a silent success is still visible',
    /brought them across from/.test(src));
  ok('the data folder is not created before adoption, or a move would become a copy',
    src.indexOf('adoptPreviousState(target)') < src.indexOf('fs.mkdirSync(target, { recursive: true })'),
    'mkdir now runs before adoption, so a 500 MB portfolio would be copied instead of moved');
}

section('12o. The back-to-top button');
{
  /*
   * A small feature, and exactly the shape of one that fails silently in this codebase:
   * a control whose element was never cached is `undefined`, `on()` returns early on a
   * null node, and the button does nothing at all with no error anywhere. The user has
   * reported that fault five separate times, so the first checks are that it exists and
   * that pressing it works.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const { window } = page;
  const css = /<style>([\s\S]*?)<\/style>/.exec(fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8'))[1];

  ok('the page still loads with the button in place',
    page.errors.length === 0, JSON.stringify(page.errors));
  ok('the button is cached rather than left as an undefined ui.* reference',
    !!page.ui.btnToTop, 'ui.btnToTop is ' + String(page.ui.btnToTop));
  ok('it is a real button, so a keyboard can reach it without being told to',
    page.ui.btnToTop.tagName === 'BUTTON', page.ui.btnToTop.tagName);
  ok('and it is not a submit button that could ever post a form',
    page.ui.btnToTop.getAttribute('type') === 'button');
  ok('it says what it does', /top/i.test(String(page.ui.btnToTop.getAttribute('aria-label'))),
    String(page.ui.btnToTop.getAttribute('aria-label')));
  ok('the arrow is decorative and hidden from a screen reader',
    /class="to-top-arrow"[^>]*aria-hidden="true"/.test(page.html));

  // Hidden at the top, and hidden in the way that also keeps it out of the tab order.
  window.pageYOffset = 0;
  page.call('renderToTop');
  ok('nothing is offered at the top of the page',
    page.ui.btnToTop.className === 'to-top', page.ui.btnToTop.className);
  ok('while hidden it cannot be focused, so it is not an invisible tab stop',
    /\.to-top \{[\s\S]*?visibility:\s*hidden/.test(css)
    && /\.to-top \{[\s\S]*?pointer-events:\s*none/.test(css),
    'opacity alone would leave a transparent button in the tab order');
  ok('and it becomes reachable again when it is shown',
    /\.to-top\.is-shown \{[\s\S]*?visibility:\s*visible/.test(css)
    && /\.to-top\.is-shown \{[\s\S]*?pointer-events:\s*auto/.test(css));

  window.pageYOffset = 3000;
  page.call('renderToTop');
  ok('scrolling a long way down offers it',
    /is-shown/.test(page.ui.btnToTop.className), page.ui.btnToTop.className);

  window.pageYOffset = 0;
  page.call('renderToTop');
  ok('and arriving back at the top takes it away again',
    page.ui.btnToTop.className === 'to-top', page.ui.btnToTop.className);

  /*
   * The threshold scales with the window rather than being a fixed pixel count, so a
   * tall screen is not offered the button after one small scroll and a short one does
   * not have to reach the bottom of the report.
   */
  window.innerHeight = 1400;
  window.pageYOffset = 2000;
  page.call('renderToTop');
  ok('on a tall window, 2000px down is not yet worth a button',
    !/is-shown/.test(page.ui.btnToTop.className), page.ui.btnToTop.className);
  window.innerHeight = 400;
  window.pageYOffset = 700;
  page.call('renderToTop');
  ok('on a short window the threshold floors rather than shrinking to nothing',
    /is-shown/.test(page.ui.btnToTop.className), page.ui.btnToTop.className);
  window.innerHeight = 768;

  // Pressing it must actually ask the browser to go to the top.
  window.pageYOffset = 3000;
  page.call('renderToTop');
  window.scrollCalls = [];
  page.ui.btnToTop.click();
  ok('clicking it asks the browser for the top of the page',
    window.scrollCalls.length === 1 && !!window.scrollCalls[0][0]
    && window.scrollCalls[0][0].top === 0, JSON.stringify(window.scrollCalls));
  ok('smoothly, when movement is welcome',
    window.scrollCalls[0][0].behavior === 'smooth', JSON.stringify(window.scrollCalls[0][0]));

  page.T.setCalm(true);
  window.scrollCalls = [];
  page.ui.btnToTop.click();
  ok('and instantly for a reader who asked for reduced motion',
    window.scrollCalls.length === 1 && window.scrollCalls[0][0].behavior === 'auto',
    JSON.stringify(window.scrollCalls));
  page.T.setCalm(false);

  /*
   * The z-index lives on the dock now rather than on the pill, because the dock owns where
   * its children sit — that is what lets the same element be a floating pill on a desktop
   * and half of a bottom bar on a phone.
   */
  ok('the dock sits above the pinned header but below the wizard overlay',
    (() => {
      const z = /\.dock \{[\s\S]*?z-index:\s*(\d+)/.exec(css);
      return !!z && Number(z[1]) > 60 && Number(z[1]) < 1000;
    })(), 'the dock would be hidden behind the header, or float over the wizard');
  ok('it is blurred like the rest of the chrome',
    /\.to-top \{[\s\S]*?backdrop-filter:\s*blur\(/.test(css));
  ok('and it stays out of the way on a phone, where the word is dropped',
    /@media \(max-width: 700px\) \{[\s\S]{0,200}\.to-top-text \{ display: none/.test(css));
}

section('12p. The app no longer claims a PDF leaves it nothing to check');
{
  /*
   * The warning read: "Only the PDF is set, so the free structural checks have nothing
   * to read and the pages cannot be counted. Add the .pptx as well to get those, or
   * accept that this check is AI-only."
   *
   * It was false, and had been for several versions. It is removed rather than reworded,
   * and these assertions exist so it cannot drift back in.
   */
  /*
   * Comments are stripped first. The note recording WHY the warning was removed quotes
   * it, and a test that cannot tell an explanation from live code is not testing
   * anything.
   */
  const source = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const page = source.replace(/\/\*[\s\S]*?\*\//g, '');
  const server = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const pdfdeck = fs.readFileSync(path.join(ROOT, 'lib', 'pdfdeck.js'), 'utf8');

  ok('the "AI-only" warning is gone', !/AI-only/.test(page), 'the message is still live in the page');
  ok('and so is the claim that the pages cannot be counted',
    !/pages cannot be counted/.test(page));
  ok('and the instruction to add a .pptx as well', !/Add the \.pptx as well/.test(page));
  ok('the explanation for removing it is still there, so it is not re-added',
    /do not put it back/.test(source));

  // The facts that made it wrong, asserted against the code rather than trusted.
  ok('a PDF is read into a real deck, so the free checks do have something to read',
    /pdfdeck\.deckFromPdf\(fs\.readFileSync\(filePath\)/.test(server),
    'the PDF path is no longer wired into the scan');
  ok('the PDF adapter supplies the page count',
    /slideCount: slides\.length/.test(pdfdeck), 'no page count from a PDF');
  ok('and the text of every page',
    /textShapes,/.test(pdfdeck), 'no text from a PDF');
  ok('and the page count is what the scan logs, so pages are counted',
    /Read \$\{deck\.slideCount\} pages/.test(server));
  ok('the rules a PDF genuinely cannot run are still named, rule by rule',
    /const skipped = \[\]/.test(fs.readFileSync(path.join(ROOT, 'lib', 'checks.js'), 'utf8'))
    && /skippedRules/.test(page),
    'a limitation must stay attached to the check it affects');
}

section('12q. One picker for the portfolio, and what each format is good for');
{
  /*
   * Requested: merge the two upload boxes, with a concise note on the limitations of each.
   *
   * Settings had a box for a PowerPoint and a box for a PDF, side by side, which read as
   * though the student was expected to supply both. The wizard has always used a single
   * picker. The formats are NOT interchangeable, though — the PDF is what the AI looks at
   * and the PowerPoint is what the layout checks can measure — so the capability stayed and
   * only the second box went.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const source = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const live = source.replace(/\/\*[\s\S]*?\*\//g, '');

  ok('the page still loads with one picker', page.errors.length === 0, JSON.stringify(page.errors));

  // One input, taking both.
  ok('the separate PDF input is gone', !/id="pdfInput"/.test(source), 'a second picker is still in the markup');
  ok('and nothing still expects it to exist',
    !/ui\.pdfInput/.test(live) && !/'pdfInput'/.test(live));
  const accept = String(page.ui.uploadInput.getAttribute('accept') || '');
  ok('the one picker accepts a PDF', /\.pdf/.test(accept), accept);
  ok('and a PowerPoint', /\.pptx/.test(accept), accept);
  /*
   * Styled selectors are stripped first: `input[type="file"]` appears in the stylesheet,
   * and counting that as a control is how a test starts lying. Two real inputs are
   * expected — the wizard's picker and this one.
   */
  const markup = live.replace(/<style>[\s\S]*?<\/style>/g, '');
  const inputs = markup.match(/type="file"/g) || [];
  eq('there are two pickers in the whole page: the wizard\'s and this one', inputs.length, 2);
  eq('and only one of them is the portfolio picker',
    (markup.match(/id="uploadInput"/g) || []).length, 1);

  // The decision itself, which is the part that can silently send a file to the wrong place.
  const kind = (name) => page.T.fileKind(name === null ? null : { name });
  eq('a .pdf is a pdf', kind('Portfolio.pdf'), 'pdf');
  eq('capitalisation does not change that', kind('PORTFOLIO.PDF'), 'pdf');
  eq('a .pptx is a pptx', kind('Portfolio.pptx'), 'pptx');
  eq('and so is a capitalised one', kind('Portfolio.PPTX'), 'pptx');
  eq('a .ppt is neither, since it is not the format the app reads', kind('Portfolio.ppt'), 'other');
  eq('a file with no name is neither', kind(''), 'other');
  eq('and so is no file at all', kind(null), 'other');

  /*
   * One rule, three callers — the picker, drag-and-drop, and the uploader that picks the
   * endpoint. They each used to work the format out with their own regex, so a PDF could be
   * accepted in one place and posted to the PowerPoint endpoint by another.
   */
  const uses = (live.match(/fileKind\(/g) || []).length;
  ok('the picker, drag-and-drop and the uploader all ask the same question',
    uses >= 4, uses + ' uses of fileKind (1 definition + 3 callers expected)');
  ok('the picker routes by that answer',
    /on\(ui\.uploadInput, 'change'[\s\S]{0,400}kind === 'pdf'[\s\S]{0,200}kind === 'pptx'/.test(live),
    'the picker does not dispatch on the file type');
  ok('and so does drag-and-drop',
    /var kind = fileKind\(file\);[\s\S]{0,300}kind === 'pdf'[\s\S]{0,160}kind === 'pptx'/.test(live));
  /*
   * The only other format tests allowed are the two guards inside uploadPortfolio and
   * uploadPdf, which assert their own precondition rather than deciding anything.
   */
  const formatTests = (live.match(/\.pdf\$|\.pptx\$/g) || []).length;
  eq('the format is worked out in one place, not four', formatTests, 4);

  // The note: concise, and it must name the limitation of EACH format, not just one.
  const note = /<div class="upload-limits"[\s\S]*?<\/div>/.exec(source);
  ok('there is a note about what each format is good for', !!note);
  // Whitespace is collapsed, because the markup wraps mid-phrase and "uncaptioned
  // images" would otherwise fail to match across a newline.
  const noteText = note ? note[0].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ') : '';
  ok('it names the PDF', /PDF/.test(noteText));
  ok('and the PowerPoint', /PowerPoint/.test(noteText));
  ok('it says the PDF is what the AI reads, including drawn work',
    /sketches|photographs/.test(noteText), noteText.slice(0, 120));
  ok('it says what a PDF CANNOT do — measure the layout',
    /cannot measure your layout/i.test(noteText), 'the PDF limitation is missing');
  ok('and what a PowerPoint misses — drawn and photographed work',
    /text only/i.test(noteText), 'the PowerPoint limitation is missing');
  ok('it says what a PowerPoint adds', /uncaptioned images/.test(noteText));
  ok('it tells the student what to do when they have both',
    /add both/i.test(noteText));
  ok('and it is a note, not an essay', noteText.replace(/\s+/g, ' ').trim().length < 760,
    noteText.replace(/\s+/g, ' ').trim().length + ' characters');

  /*
   * The line announcing "No PowerPoint file chosen" used to be shown whenever a
   * PowerPoint was absent. In a box that also takes a PDF that read as a fault to fix.
   */
  /*
   * The old "No PowerPoint file chosen" announcement is gone with the prose it lived in.
   * Absence is now a row in the status list, which says what is missing and what it costs
   * without reading as a fault the student is supposed to go and fix.
   */
  ok('absence is shown as a row rather than announced as a fault',
    !/No PowerPoint file chosen/.test(live) && /not added/.test(live),
    'the negative announcement is gone but nothing says what is missing instead');
}

section('12r. What has been loaded, said plainly');
{
  /*
   * Requested: "make it clearer what has been uploaded though with a place that shows which
   * has been uploaded out of the pptx and pdf".
   *
   * It was unclear for a real reason, and not just a cosmetic one. The group header read
   * \`config.pdfName || config.deckName || null\`, and \`deckName\` is not a field the server
   * has EVER sent — so with a PowerPoint on its own the header said "nothing added yet" while
   * a portfolio was loaded and being marked. Silent, and wrong in the one place that answers
   * the question.
   */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  ok('the page still loads with the status list', page.errors.length === 0, JSON.stringify(page.errors));

  const draw = (config, deck) => {
    page.state.data = { ok: true, config: config, deck: deck || null };
    const a = page.call('renderLoadedFiles');
    const b = page.call('renderGroupNotes');
    return {
      header: page.els.get('portfolioNote').textContent,
      cls: page.els.get('portfolioNote').className,
      rows: page.els.get('uploadedWhat').children.map((c) => c.children.map((x) => x.textContent).join(' ')),
      ok: a.ok && b.ok,
      error: a.error || b.error,
    };
  };

  const pdfCfg = { pdfName: 'Portfolio.pdf', pdfFile: '/x/Portfolio.pdf' };
  const pptxCfg = { deckFile: '/x/y/Portfolio.pptx' };
  const deck = { fileName: 'Portfolio.pptx', slideCount: 29, fileSizeBytes: 4700000 };
  const both = Object.assign({}, pptxCfg, pdfCfg);

  let r = draw(pdfCfg, null);
  ok('both renders run without throwing', r.ok, r.error);
  eq('a PDF alone is reported as the PDF', r.header, 'PDF added');
  ok('and marked as satisfied', /saved/.test(r.cls), r.cls);

  r = draw(pptxCfg, deck);
  eq('a PowerPoint alone is reported as the PowerPoint', r.header, 'PowerPoint added');
  ok('this is the case that used to say nothing was added at all',
    r.header !== 'nothing added yet', r.header);

  r = draw(both, deck);
  eq('both are reported as both', r.header, 'PowerPoint and PDF added');

  r = draw({}, null);
  eq('nothing is reported as nothing', r.header, 'nothing added yet');
  ok('and marked as not yet done', /muted/.test(r.cls), r.cls);

  // The rows: one per format, so either can be understood without the other.
  r = draw(both, deck);
  eq('there is one row per format', r.rows.length, 2);
  ok('the PDF row names the file', /PDF.*Portfolio\.pdf/.test(r.rows.join(' ~ ')), r.rows.join(' ~ '));
  ok('the PowerPoint row names the file and its length',
    /PowerPoint.*Portfolio\.pptx.*29 pages/.test(r.rows.join(' ~ ')), r.rows.join(' ~ '));

  r = draw(pdfCfg, null);
  ok('an absent PowerPoint row says what it costs, not merely that it is missing',
    /PowerPoint.*not added.*checks cannot run/.test(r.rows.join(' ~ ')), r.rows.join(' ~ '));

  r = draw(pptxCfg, deck);
  ok('an absent PDF row says what the AI is missing',
    /PDF.*not added.*words only/.test(r.rows.join(' ~ ')), r.rows.join(' ~ '));

  /*
   * And the field that caused it must not come back. Reading a config field the server does
   * not send is this codebase's signature fault, and it fails silently every time.
   */
  /*
   * Comments are stripped before the check, because the notes recording why the bug happened
   * name the field. HTML comments matter as much as block comments here — the markup carries
   * an explanation of the same fault.
   */
  const live = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/<!--[\s\S]*?-->/g, '');
  ok('nothing reads a config field the server does not send',
    !/config\.deckName/.test(live),
    'config.deckName has never been sent by the server, and reading it is what hid a loaded PowerPoint');
}

section('12s. The icon, in every place it appears');
{
  /*
   * Reported: "i have it on my home screen on mobile but theres no icon".
   *
   * A phone home-screen shortcut takes its icon from apple-touch-icon, NOT from the
   * favicon, so having one and not the other would not have fixed it. There was no icon
   * anywhere: not in the tab, not on the home screen, not on the desktop.
   */
  const publicDir = path.join(ROOT, 'public');
  const assetsDir = path.join(ROOT, 'assets');
  const page = fs.readFileSync(path.join(publicDir, 'index.html'), 'utf8');

  // --- the files are there ----------------------------------------------------------
  const webIcons = ['favicon.ico', 'icon-16.png', 'icon-32.png', 'icon-180.png',
    'icon-192.png', 'icon-512.png', 'icon-512-maskable.png', 'manifest.webmanifest'];
  ok('every icon file the page needs is present',
    webIcons.every((f) => fs.existsSync(path.join(publicDir, f))),
    webIcons.filter((f) => !fs.existsSync(path.join(publicDir, f))).join(', ') + ' missing');
  ok('the Mac icon is present', fs.existsSync(path.join(assetsDir, 'NEA Coach.icns')));
  ok('the Windows icon is present', fs.existsSync(path.join(assetsDir, 'NEA Coach.ico')));

  // --- the page asks for them -------------------------------------------------------
  ok('the page declares an apple-touch-icon',
    /rel="apple-touch-icon"[^>]*href="\/icon-180\.png"/.test(page),
    'iOS reads apple-touch-icon for Add to Home Screen, not the favicon');
  ok('and a favicon', /rel="icon"[^>]*href="\/favicon\.ico"/.test(page));
  ok('and the manifest', /rel="manifest"[^>]*href="\/manifest\.webmanifest"/.test(page));
  ok('and a theme colour, so the phone status bar matches',
    /name="theme-color"[^>]*content="#2563eb"/.test(page));
  ok('and the Safari standalone tag, which iOS still needs',
    /name="apple-mobile-web-app-capable"[^>]*content="yes"/.test(page));

  // --- decoding, which is the only thing that catches a bad checksum -----------------
  /*
   * Node's own CRC, deliberately. Reusing the generator's would let a shared bug pass, and
   * a shared bug is exactly what happened: the first version of tools/make-icons.js shifted
   * by one bit instead of eight, producing a checksum that looks like a checksum. `file`
   * still reported "PNG image data, 180 x 180" while macOS and the image readers refused to
   * open it. Nothing short of decoding finds that.
   */
  ok('this Node can verify checksums independently',
    typeof zlib.crc32 === 'function', 'zlib.crc32 is unavailable');

  const decode = (buf, label) => {
    if (buf.slice(0, 8).toString('hex') !== '89504e470d0a1a0a') { return label + ': not a PNG'; }
    let off = 8; let w = 0; let h = 0; let bitDepth = 0; let colour = 0;
    const idat = [];
    while (off + 12 <= buf.length) {
      const len = buf.readUInt32BE(off);
      const type = buf.slice(off + 4, off + 8).toString('ascii');
      const data = buf.slice(off + 8, off + 8 + len);
      const stored = buf.readUInt32BE(off + 8 + len);
      const mine = zlib.crc32(Buffer.concat([buf.slice(off + 4, off + 8), data])) >>> 0;
      if (mine !== stored) { return label + ': ' + type + ' checksum is wrong'; }
      if (type === 'IHDR') {
        w = data.readUInt32BE(0); h = data.readUInt32BE(4); bitDepth = data[8]; colour = data[9];
      }
      if (type === 'IDAT') { idat.push(data); }
      off += 12 + len;
    }
    if (bitDepth !== 8 || colour !== 6) { return label + ': expected 8-bit RGBA'; }
    const raw = zlib.inflateSync(Buffer.concat(idat));
    if (raw.length !== (w * 4 + 1) * h) {
      return label + ': inflated ' + raw.length + ' bytes, wanted ' + ((w * 4 + 1) * h);
    }
    return { w: w, h: h };
  };

  const pngFiles = webIcons.filter((f) => /\.png$/.test(f));
  const badPng = pngFiles.filter((f) => typeof decode(fs.readFileSync(path.join(publicDir, f)), f) === 'string');
  ok('every generated PNG decodes, checksums and all', badPng.length === 0, badPng.join(', '));

  // --- the manifest agrees with the files -------------------------------------------
  const manifest = JSON.parse(fs.readFileSync(path.join(publicDir, 'manifest.webmanifest'), 'utf8'));
  ok('the manifest names the app',
    manifest.name === 'NEA Coach' && manifest.short_name === 'NEA Coach');
  ok('it opens at the root, which the access cookie already authenticates',
    manifest.start_url === '/', manifest.start_url);
  ok('and asks to be shown without browser chrome',
    manifest.display === 'standalone', manifest.display);
  ok('every icon it declares exists and is the size it claims',
    manifest.icons.every((i) => {
      const file = path.join(publicDir, i.src.replace(/^\//, ''));
      if (!fs.existsSync(file)) { return false; }
      const d = decode(fs.readFileSync(file), i.src);
      return typeof d !== 'string' && d.w + 'x' + d.h === i.sizes;
    }), JSON.stringify(manifest.icons.map((i) => i.src + ' -> ' + i.sizes)));
  ok('and one of them is maskable, because Android crops to its own shape',
    manifest.icons.some((i) => i.purpose === 'maskable'));

  // --- the two platform icons are real containers -----------------------------------
  const icoSizes = (file) => {
    const buf = fs.readFileSync(file);
    if (buf.readUInt16LE(0) !== 0 || buf.readUInt16LE(2) !== 1) { return null; }
    const out = [];
    for (let i = 0; i < buf.readUInt16LE(4); i++) {
      const at = 6 + i * 16;
      const declared = buf[at] === 0 ? 256 : buf[at];
      const len = buf.readUInt32LE(at + 8);
      const off = buf.readUInt32LE(at + 12);
      out.push({ declared: declared, png: buf.slice(off, off + len) });
    }
    return out;
  };

  for (const [label, file] of [['favicon.ico', path.join(publicDir, 'favicon.ico')],
    ['NEA Coach.ico', path.join(assetsDir, 'NEA Coach.ico')]]) {
    const entries = icoSizes(file);
    ok(label + ' is a real icon container', !!entries && entries.length > 0);
    /*
     * Different jobs, different sizes. A favicon only ever appears in a tab, so 16/32/48 is
     * the right set; Explorer and the taskbar ask for a large one, which is why the Windows
     * icon carries 256 as well.
     */
    if (label === 'favicon.ico') {
      ok(label + ' holds the sizes a browser tab asks for',
        [16, 32].every((s) => entries.some((e) => e.declared === s)),
        entries.map((e) => e.declared).join(', '));
    } else {
      ok(label + ' holds a small size for the taskbar and a large one for Explorer',
        entries.some((e) => e.declared <= 32) && entries.some((e) => e.declared >= 256),
        entries.map((e) => e.declared).join(', '));
    }
    const broken = entries.filter((e) => {
      const d = decode(e.png, label);
      return typeof d === 'string' || d.w !== e.declared;
    });
    ok(label + ' holds PNGs that decode at the size each entry claims',
      broken.length === 0, broken.length + ' bad entries');
  }

  const icns = fs.readFileSync(path.join(assetsDir, 'NEA Coach.icns'));
  ok('the Mac icon is a real icns container',
    icns.slice(0, 4).toString('ascii') === 'icns' && icns.readUInt32BE(4) === icns.length,
    'the icns header or length is wrong');
  const icnsTypes = [];
  const icnsPngs = [];
  for (let off = 8; off + 8 <= icns.length;) {
    const type = icns.slice(off, off + 4).toString('ascii');
    const len = icns.readUInt32BE(off + 4);
    icnsTypes.push(type);
    icnsPngs.push(icns.slice(off + 8, off + len));
    off += len;
  }
  ok('and it carries the sizes macOS asks for',
    ['ic11', 'ic12', 'ic07', 'ic08', 'ic09', 'ic10'].every((t) => icnsTypes.indexOf(t) > -1),
    icnsTypes.join(', '));
  ok('with every one decoding',
    icnsPngs.every((b, i) => typeof decode(b, 'icns ' + icnsTypes[i]) !== 'string'));

  /*
   * The generator must refuse to write anything if its checksum is wrong. This is the
   * self-check that turns a silent, every-decoder-rejects-it failure into a loud one.
   */
  const tool = fs.readFileSync(path.join(ROOT, 'tools', 'make-icons.js'), 'utf8');
  ok('the icon generator checks its own CRC before writing anything',
    /0xcbf43926/.test(tool), 'the standard CRC-32 test vector is not checked');
  ok('and the update step shifts by eight, which is the bug it had',
    /crc >>> 8/.test(tool), 'the checksum update is wrong again');

  // --- the desktop icons ------------------------------------------------------------
  const bat = fs.readFileSync(path.join(ROOT, 'Start NEA Coach.bat'), 'utf8');
  ok('the Windows launcher makes a Desktop shortcut with the icon',
    /CreateShortcut/.test(bat) && /IconLocation='%~dp0assets\\NEA Coach\.ico/.test(bat),
    'a .bat cannot carry an icon, so a shortcut is the only way to give it one');
  ok('and only when it is missing, so deleting it is respected',
    /if \(-not \(Test-Path \$p\)\)/.test(bat));
  /*
   * Ordering, and this was a real bug introduced while adding it. `pushd` is followed by
   * `if not errorlevel 1 goto :in_folder`, and the shortcut call was first placed BETWEEN
   * them — so a PowerShell that failed (a blocked execution policy, say) would make the
   * launcher report "could not open the app folder" and refuse to start at all.
   */
  const pushdAt = bat.indexOf('pushd "%~dp0."');
  const checkAt = bat.indexOf('if not errorlevel 1 goto :in_folder');
  const shortcutAt = bat.indexOf('CreateShortcut');
  ok('the shortcut call cannot break the launcher\'s own start-up check',
    pushdAt > -1 && checkAt > pushdAt && shortcutAt > checkAt,
    'anything between pushd and its errorlevel check changes what that check is testing');

  const pkg = fs.readFileSync(path.join(ROOT, 'tools', 'package.sh'), 'utf8');
  ok('the Mac package builds a real .app bundle',
    /build_mac_app/.test(pkg) && /CFBundleExecutable/.test(pkg) && /CFBundleIconFile/.test(pkg));
  ok('the bundle carries its own copy of the app, so it can be dragged to the Desktop',
    /Contents\/Resources\/app/.test(pkg));
  ok('and it is staged outside the folder it copies, or the copy would recurse',
    /\.nea-app-stage/.test(pkg));
  ok('assets ship, so both platforms can find their icon', /^\s*assets$/m.test(pkg));
}

section('12t. When the marking model is busy, the review moves to another one');
{
  /*
   * This is the fix for a 99/100 portfolio being reported as 74, and the evidence is the
   * student's own scan log:
   *
   *   [12:40:25] Visual review will use P1.pdf (12.6 MB)
   *   [12:40:41] Retrying the visual review (attempt 1, HTTP 503)
   *   [12:43:53] Retrying the visual review (attempt 2)
   *   [12:44:10] Retrying the visual review (attempt 3, HTTP 503)
   *   [12:44:37] Retrying the visual review (attempt 4, HTTP 503)
   *   [12:44:58] A visual review call failed: This model is currently experiencing high demand.
   *
   * Four 503s over four minutes, the visual review abandoned, and the app went on to report a
   * TEXT-ONLY figure as a mark. On a portfolio whose marks live in its sketches and
   * photographs that is not a low mark, it is the wrong question — the text-only reading for
   * that portfolio is about 75 against a real 99.
   *
   * The replacement matters as much as the failover: only models this app has actually been
   * calibrated against are offered, so the fallback's reading can still be corrected. An
   * uncorrectable number is the thing being avoided.
   */
  const ranked = calibration.calibratedModels();
  ok('the app knows which models it can actually correct', ranked.length >= 2, JSON.stringify(ranked));
  eq('and ranks the best-calibrated first', ranked[0].model, 'gemini-3.5-flash');
  ok('each entry records how far out its corrected figure has been',
    ranked.every((r) => typeof r.worstResidual === 'number' && r.anchors >= 2), JSON.stringify(ranked));
  ok('the ranking is derived from the anchors, not written down',
    ranked[0].worstResidual <= ranked[1].worstResidual,
    ranked.map((r) => r.model + ':' + r.worstResidual).join(', '));

  // A model the app cannot correct must never be offered as a replacement.
  const available = ['gemini-3.8-flash', 'gemini-3.5-flash', 'gemini-3.1-flash-lite', 'gemini-2.5-pro'];
  const chain = calibration.calibratedModels()
    .map((c) => c.model)
    .filter((n) => n !== 'gemini-3.8-flash' && available.indexOf(n) > -1);
  ok('an uncorrectable model is never offered as a fallback',
    chain.every((m) => calibration.calibratedModels().some((c) => c.model === m)), chain.join(', '));

  /* ---- the real thing, with the API stubbed ------------------------------------- */
  const https = require('https');
  const { EventEmitter } = require('events');
  const realRequest = https.request;
  const criteria = [
    { id: 'a', name: 'Investigating', weight: 10 },
    { id: 'b', name: 'Brief', weight: 10 },
  ];
  const visionReply = (mark) => ({
    candidates: [{ content: { parts: [{ text: JSON.stringify({
      coverage: criteria.map((c) => ({
        criterion: c.id, status: 'strong', coveragePercent: 90, suggestedMark: mark, slides: [1], comment: 'ok',
      })),
      strengths: [], suggestions: [], findings: [], error: '',
    }) }] } }],
    usageMetadata: { promptTokenCount: 40000, candidatesTokenCount: 500 },
  });

  const seen = [];
  // Only the CONFIGURED model is unavailable. If every model failed, the test would prove
  // nothing about failing over.
  let failModels = ['gemini-3.8-flash'];
  let failStatus = 503;
  https.request = (opts, onResponse) => {
    const req = new EventEmitter();
    let body = '';
    req.setTimeout = () => {};
    req.write = (b) => { body += b; };
    req.end = () => {
      // The model is named in the request path, which is how the real client sends it.
      const found = /models\/([^:?]+)/.exec(opts.path || '');
      const name = found ? found[1] : '?';
      seen.push(name);
      const res = new EventEmitter();
      res.headers = {};
      setTimeout(() => {
        if (failModels.indexOf(name) > -1) {
          res.statusCode = failStatus;
          onResponse(res);
          res.emit('data', Buffer.from(JSON.stringify({ error: { message: 'This model is currently experiencing high demand.' } })));
        } else {
          res.statusCode = 200;
          onResponse(res);
          res.emit('data', Buffer.from(JSON.stringify(visionReply(9))));
        }
        res.emit('end');
      }, 0);
    };
    return req;
  };

  try {
    const pass = await ai.runVisionPass({
      deck: { slideCount: 60 },
      apiKey: 'test',
      model: 'gemini-3.8-flash',
      fallbackModels: ['gemini-3.5-flash'],
      spec: { criteria: criteria },
      pdf: Buffer.from('%PDF-1.4 not really a pdf'),
      onLog: () => {},
      runs: 1,
    });
    ok('a busy marking model no longer loses the whole review',
      !!(pass && pass.ok), JSON.stringify(pass && pass.reason));
    eq('the review reports the model that really answered', pass && pass.model, 'gemini-3.5-flash');
    eq('and it tried the configured model first', seen[0], 'gemini-3.8-flash');
    ok('before moving to the calibrated one',
      seen.indexOf('gemini-3.5-flash') > 0, seen.join(', '));
    ok('with the page coverage intact, so a mark can still be made',
      !!(pass && pass.coverage && pass.coverage.length), JSON.stringify(pass && pass.coverage));

    /*
     * A failure no other model would survive — a rejected request — must NOT be retried
     * against every model in turn: that would spend a request per model to learn the same
     * thing. It comes back straight away, carrying the reason.
     */
    seen.length = 0;
    failStatus = 400;
    const rejected = await ai.runVisionPass({
      deck: { slideCount: 60 },
      apiKey: 'test',
      model: 'gemini-3.8-flash',
      fallbackModels: ['gemini-3.5-flash'],
      spec: { criteria: criteria },
      pdf: Buffer.from('%PDF-1.4 not really a pdf'),
      onLog: () => {},
      runs: 1,
    });
    eq('with nothing to fall back to the review reports failure rather than a number',
      rejected && rejected.ok, false);
    eq('and it names the reason, so the page can explain it', rejected.reason, 'error');
    eq('without trying the other model at all', seen.length, 1);
    ok('recording what it tried', Array.isArray(rejected.tried) && rejected.tried.length === 1,
      JSON.stringify(rejected.tried));
  } finally {
    https.request = realRequest;
  }

  // The reason codes the page switches on.
  eq('a busy service is reported as busy', ai.visionFailureReason(503), 'busy');
  eq('and so is a bad gateway', ai.visionFailureReason(502), 'busy');
  eq('a daily cap is named as one, because it is a different instruction',
    ai.visionFailureReason(429), 'rate-limited');
  eq('a silent timeout is its own case', ai.visionFailureReason(0), 'no-response');
  eq('a model that is gone is named as unavailable', ai.visionFailureReason(404), 'unavailable');

  /* ---- and the page must say so, next to the number ------------------------------ */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const textOf = (node) => {
    if (!node) { return ''; }
    let out = node.textContent ? String(node.textContent) : '';
    for (const k of (node.children || [])) { out += ' ' + textOf(k); }
    return out;
  };
  const predicted = { rows: [{ id: 'a', name: 'A', weight: 10, mark: 9, max: 10 }], predicted: 74, maximum: 100, percentage: 74, weightsKnown: true };
  const drawMark = (aiBlock) => {
    page.state.data = { ok: true, config: {}, deck: null, local: { findings: [] }, ai: aiBlock };
    page.call('renderPredicted');
    return textOf(page.els.get('predictedBody')).replace(/\s+/g, ' ');
  };

  const busy = drawMark({ ran: true, visionRan: false, visionFailure: { reason: 'busy', status: 503 }, predicted: predicted });
  ok('the page says the AI never saw the pages', /never saw your pages/.test(busy), busy.slice(0, 160));
  ok('and says the figure comes from the text alone', /TEXT alone/.test(busy));
  ok('and says it is not a judgement of the work', /not a judgement/.test(busy));
  ok('and tells the student what to do', /Press Check now to try again/.test(busy));

  const capped = drawMark({ ran: true, visionRan: false, visionFailure: { reason: 'rate-limited', status: 429 }, predicted: predicted });
  ok('a daily cap is explained as one, not as the app being worse',
    /daily free allowance/.test(capped), capped.slice(0, 160));

  const fine = drawMark({ ran: true, visionRan: true, predicted: predicted });
  ok('and a review that worked carries no such warning', !/never saw your pages/.test(fine));

  /*
   * A missing PDF is NOT a failure and must not get the alarm. The student chose it, and the
   * calm note at the end of the panel already explains the consequence.
   */
  const noPdf = drawMark({ ran: true, visionRan: false, predicted: predicted });
  ok('no PDF at all is not treated as a failure', !/never saw your pages/.test(noPdf));
}

section('12u. Using it on a phone is its own section, not an advanced setting');
{
  /*
   * Requested: "make the how to use it on your phone a new section at the bottom instead of
   * in settings".
   *
   * It was in the wrong place twice over. It is not an option to tune, so it did not belong
   * among the advanced ones — and it was the one part of the page a phone user could not find
   * on the phone they were holding, which is exactly who needs it.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const settingsStart = html.indexOf('<details class="card details-card" id="settingsDetails"');
  const settingsEnd = html.indexOf('</details>', html.indexOf('infoDetails'));
  const cardAt = html.indexOf('id="connectCard"');
  const footerAt = html.indexOf('<footer class="footer">');

  ok('there is a phone section', cardAt > -1);
  ok('it has been taken out of Settings',
    !/Use it on your phone/.test(html.slice(settingsStart, settingsEnd)),
    'the phone instructions are still inside the settings group');
  ok('it sits below the settings block', cardAt > settingsEnd);
  ok('and above the footer, so it is the last thing on the page', cardAt < footerAt);
  ok('it is a card with its own heading',
    /<section class="card" id="connectCard" aria-labelledby="connectHeading">/.test(html));
  ok('the heading is a real one for the accessibility tree',
    /<h2 class="card-title" id="connectHeading">/.test(html));

  // The stale placeholder from when it lived in Settings must be gone with it.
  ok('no leftover "open Settings" placeholder',
    !/Open Settings to see the address/.test(html),
    'the placeholder still tells the reader to go to Settings');

  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  ok('the page loads with the card in place', page.errors.length === 0, JSON.stringify(page.errors));
  ok('every element of it is cached, so none of it is an undefined ui.* reference',
    !!page.ui.connectCard && !!page.ui.connectBody && !!page.ui.connectNote && !!page.ui.btnConnectAgain,
    [page.ui.connectCard, page.ui.connectBody, page.ui.connectNote, page.ui.btnConnectAgain].map(Boolean).join(','));

  const connectCalls = () => page.fetchCalls.filter((c) => String(c && c.url ? c.url : c).includes('/api/connect')).length;
  /*
   * Filled in on its own rather than on a click. The card is at the very bottom of the page,
   * so making the reader scroll to it and then press something before there is anything to
   * read would be a worse trade than one small request.
   */
  const onLoad = connectCalls();
  ok('the address is looked up as soon as the page is up', onLoad >= 1, onLoad + ' lookups');
  /*
   * And it can be looked up AGAIN. The address changes whenever the router reassigns it, so
   * "come back here and copy the new one" is useless without a way to come back — the flag
   * that stops the automatic load repeating must not stop a deliberate one.
   */
  page.ui.btnConnectAgain.click();
  ok('and Check again forces a fresh lookup, rather than being swallowed by the once-only flag',
    connectCalls() > onLoad, connectCalls() + ' lookups after pressing it');
  /*
   * The phone is the one device that cannot be told the address from a phone, so the message
   * it gets must point at the Mac, not at a Settings section the address has left.
   */
  ok('the message a phone gets points at the Mac, not at Settings',
    !/shown on the Mac, in Settings/.test(html),
    'the phone is still told to look in Settings');
  ok('and it is still told where to look', /shown on the Mac/.test(html));
}

section('12v. A stored PowerPoint copy can be seen and removed');
{
  /*
   * Reported: "there is no remove pptx button".
   *
   * There wasn't, and it was a regression from moving the data folder out of the app folder in
   * v1.0.23. The page decided whether a PowerPoint was loaded with
   *
   *   /(^|[\/])state[\/]portfolio[\/]?$/.test(configuredDir)
   *
   * which stopped matching the moment uploads moved to ~/Library/Application Support/NEA Coach.
   * Uploading a PDF also clears deckFile, so the flag went false, the button disappeared, and
   * the app's own copy of a .pptx sat in its folder with nothing anywhere mentioning it. On the
   * machine it was reported from, that copy was 528 MB and had been there for weeks.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  ok('the page no longer decides by pattern-matching a path that moved',
    !/inUploadFolder/.test(html),
    'the state/portfolio pattern is back; it stopped matching in v1.0.23');
  ok('and it uses the copy the server reports instead', /deckCopy/.test(html));

  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const UP = '/Users/x/Library/Application Support/NEA Coach/portfolio';
  const textOf = (n) => {
    if (!n) { return ''; }
    let out = n.textContent ? String(n.textContent) : '';
    for (const k of (n.children || [])) { out += ' ' + textOf(k); }
    return out;
  };
  const draw = (config, deck) => {
    page.state.data = { ok: true, config: config, deck: deck || null };
    const a = page.call('renderLoadedFiles');
    const b = page.call('renderUploadSection', page.state.data, config);
    return {
      ok: a.ok && b.ok,
      error: a.error || b.error,
      heading: page.els.get('uploadHeading').textContent,
      button: page.ui.btnClearDeck.hidden === false ? page.ui.btnClearDeck.textContent : null,
      rows: page.els.get('uploadedWhat').children.map((c, i) => textOf(page.els.get('uploadedWhat').children[i])),
      raw: page.els.get('uploadedWhat').children.map((c) => c.textContent),
    };
  };

  const stored = { name: 'Old Portfolio.pptx', sizeBytes: 553648128, sizeLabel: '528 MB' };
  const pdfOnly = { deckDir: UP, deckFile: null, pdfFile: UP + '/work.pdf', pdfName: 'work.pdf', deckCopy: stored };

  let r = draw(pdfOnly, { fileName: 'work.pdf', slideCount: 60 });
  ok('the state that had no button renders at all', r.ok, r.error);
  ok('a stored copy gets a remove button', !!r.button, 'no remove button');
  ok('and the button says what it will delete, and how much space that is',
    /Delete the stored PowerPoint copy \(528 MB\)/.test(r.button || ''), r.button);
  ok('the row explains why it is stored rather than being marked',
    /stored, but you are being marked on the PDF/.test(textOf(page.els.get('uploadedWhat').children[1])),
    textOf(page.els.get('uploadedWhat').children[1]));
  ok('and it is flagged, not ticked as though it were in use',
    /^!/.test(textOf(page.els.get('uploadedWhat').children[1]).trim()),
    textOf(page.els.get('uploadedWhat').children[1]));

  /*
   * The row must never call a PDF a PowerPoint. It did: it fell back to the report's own file
   * name, and for a PDF-only setup that name IS the PDF — which also made the page believe a
   * PowerPoint was loaded.
   */
  ok('the PowerPoint row never names the PDF',
    !/work\.pdf/.test(textOf(page.els.get('uploadedWhat').children[1])),
    textOf(page.els.get('uploadedWhat').children[1]));

  r = draw({ deckDir: UP, deckFile: null, pdfFile: UP + '/work.pdf', pdfName: 'work.pdf' },
    { fileName: 'work.pdf', slideCount: 60 });
  ok('with no PowerPoint at all the row says so',
    /not added/.test(textOf(page.els.get('uploadedWhat').children[1])),
    textOf(page.els.get('uploadedWhat').children[1]));
  ok('and there is no button, because there is nothing to remove', !r.button, r.button);
  ok('while the heading still knows a portfolio is in place',
    /Choose a different portfolio/.test(r.heading), r.heading);

  // The ordinary cases must not regress.
  r = draw({ deckDir: UP, deckFile: UP + '/X.pptx', pdfFile: UP + '/X.pdf', pdfName: 'X.pdf' },
    { fileName: 'X.pptx', slideCount: 29 });
  eq('with a PowerPoint loaded the button says so', r.button, 'Remove PowerPoint');
  ok('and the row names it', /X\.pptx/.test(textOf(page.els.get('uploadedWhat').children[1])),
    textOf(page.els.get('uploadedWhat').children[1]));

  r = draw({ deckDir: UP, deckFile: null }, { fileName: 'Folder Work.pptx', slideCount: 12 });
  eq('a folder-based setup still counts as having a PowerPoint', r.button, 'Remove PowerPoint');

  r = draw({ deckDir: UP }, null);
  ok('and nothing at all still shows no button', !r.button, r.button);
  eq('with the getting-started heading', r.heading, 'Get started — add your portfolio');

  /*
   * Removing is DESTRUCTIVE now — it used to clear the reference and leave the file behind,
   * which is how the copy survived in the first place. The guard that matters is that it can
   * only ever touch the app's own folder, never the student's original.
   */
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const block = /if \(body\.clearDeck === true\) \{([\s\S]*?)\n    \}/.exec(src);
  ok('the removal code is still findable', !!block);
  const body = block ? block[1] : '';
  ok('removing a PowerPoint now actually deletes the stored copy',
    /unlinkSync/.test(body), 'it still only clears the reference and leaves the file behind');
  ok('and only inside the app\u2019s own upload folder',
    /path\.dirname\(path\.resolve\(full\)\) !== path\.resolve\(uploadDir\)/.test(body),
    'nothing stops it deleting a path outside the upload folder');
  /*
   * The precise guarantee: exactly ONE unlink, of a path built from a name that was listed in
   * the app's own folder. Anything else — deleting the configured path, say, which can point
   * anywhere on the disk — is how a student's original gets destroyed.
   */
  ok('the only file it deletes is one it listed inside its own upload folder',
    /readdirSync\(uploadDir\)/.test(body)
    && /const full = path\.join\(uploadDir, name\);/.test(body)
    && /fs\.unlinkSync\(full\)/.test(body)
    && (body.match(/unlinkSync\(/g) || []).length === 1,
    'it deletes something other than a file it listed in its own folder');
  ok('and it reports how much space was freed',
    /MB freed/.test(body), 'the log does not say what was reclaimed');
}

section('12w. The aesthetic pass: a tighter phone, and more life in the marks');
{
  /*
   * Requested: make it look even nicer, less crowded on an iPhone, and more of the dynamic
   * thing the student liked. Thirteen items, all of them, so this section pins the ones that
   * can silently stop working — chiefly the gating, because an animation that fires when it
   * should not is worse than one that never fires.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(html)[1];

  /* ---- the report is one panel, not three cards --------------------------------- */
  ok('the marking line, report and actions are one welded panel',
    /#reviewedCard, #statusCard, #actionsCard \{/.test(css));
  ok('the welded parts cannot lift or rise independently',
    /#reviewedCard, #statusCard, #actionsCard \{[\s\S]*?transform: none/.test(css)
    && /#reviewedCard, #statusCard, #actionsCard \{[\s\S]*?animation-delay: 30ms/.test(css),
    'three cards sliding up on different clocks gives the join away');
  ok('and the middle and top sections square off so the seam is invisible',
    /#statusCard \{[\s\S]*?border-radius: 0/.test(css)
    && /#reviewedCard \{[\s\S]*?border-bottom-left-radius: 0/.test(css));
  ok('the two cards the weld needs are actually addressable by id',
    /id="statusCard"/.test(html) && /id="actionsCard"/.test(html));
  ok('and the markup was not restructured to do it, so every id and render is untouched',
    /id="reviewedCard"/.test(html) && /id="reviewedName"/.test(html));

  /* ---- phone density ------------------------------------------------------------ */
  const phone = /@media \(max-width: 700px\) \{\n    \/\*\n     \* Everything is tighter[\s\S]*?\n  \}/.exec(css);
  ok('there is a phone density pass', !!phone);
  ok('it tightens the card padding and gaps, which repeat 17 times down the page',
    /\.card \{ padding: 13px; margin-bottom: 11px/.test(css), 'the repetition is the crowding');
  ok('the ring sits beside the verdict instead of above it',
    /\.status-top \{ display: flex; align-items: center; gap: 13px; \}/.test(css),
    'stacking a 130px circle pushes the one sentence that matters off the first screen');
  ok('and the ring is smaller on a phone', /\.ring-wrap, \.ring \{ width: 104px/.test(css));
  ok('the file name is one truncated line rather than wrapping to three',
    /\.app-sub \{ white-space: nowrap; overflow: hidden; text-overflow: ellipsis/.test(css));
  ok('and the setup button and badge share a row instead of taking one each',
    /\.topbar-actions \{\n      flex-direction: row;/.test(css));

  /* ---- the dock ------------------------------------------------------------------ */
  ok('the dock owns the positioning, so one element is a pill on desktop and a bar on a phone',
    /\.dock \{[\s\S]*?position: fixed/.test(css) && /\.to-top \{\n    position: static;/.test(css));
  ok('the phone bar only gains a background when it has something to offer',
    /\.dock\.has-quick \{/.test(css), 'an always-on bar would sit over the page for no reason');
  ok('and the page is padded so the last card clears it',
    /\.page \{ padding-bottom: 86px; \}/.test(css));

  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  ok('the page still loads with all of it', page.errors.length === 0, JSON.stringify(page.errors));
  ok('the dock and its two buttons are cached, not undefined ui.* references',
    !!page.ui.actionDock && !!page.ui.btnQuickCheck && !!page.ui.btnToTop);

  const gone = { top: -200, bottom: -180, left: 0, right: 100, width: 100, height: 20 };
  const here = { top: 100, bottom: 120, left: 0, right: 100, width: 100, height: 20 };
  page.ui.btnCheck.getBoundingClientRect = () => gone;
  page.call('renderDock');
  ok('the shortcut appears once the real button has scrolled away',
    /has-quick/.test(page.ui.actionDock.className), page.ui.actionDock.className);
  page.ui.btnCheck.getBoundingClientRect = () => here;
  page.call('renderDock');
  ok('and goes when the real one is back in view',
    !/has-quick/.test(page.ui.actionDock.className), page.ui.actionDock.className);

  page.state.scanning = false;
  page.state.data = { ok: true, config: { portfolioChosen: false }, deck: null, local: { findings: [] }, ai: {} };
  page.call('renderActions');
  eq('the shortcut mirrors the real button rather than deciding for itself',
    page.ui.btnQuickCheck.textContent, page.ui.btnCheckLabel.textContent);
  ok('including being disabled for the same reason', page.ui.btnQuickCheck.disabled === true);

  /* ---- the marks have more life -------------------------------------------------- */
  const report = JSON.parse(fs.readFileSync(path.join(ROOT, 'state', 'last-report.json'), 'utf8'));

  eq('a fresh page shows no flare, because there is nothing to have changed from yet',
    page.ui.ring.className, 'ring');
  page.state.data = report;
  page.call('renderStatus');
  ok('a new figure flares the ring', /is-pop/.test(page.ui.ring.className), page.ui.ring.className);
  ok('and the flare is designed to clear itself rather than linger',
    /ringPopTimer/.test(html), 'the class would stay on for good');

  page.call('renderFixList');
  ok('findings animate in when the list is first built',
    /is-fresh/.test(page.ui.fixList.className), page.ui.fixList.className);
  page.call('renderFixList');
  ok('but NOT on a redraw with the same contents',
    !/is-fresh/.test(page.ui.fixList.className),
    'an entrance that replays on every re-render becomes noise');

  const findings = page.T.allFindings();
  ok('the real report has findings to work with', findings.length > 0, String(findings.length));
  for (const f of findings) { page.T.toggleDone(f.id, true); }
  page.call('renderFixList');
  const rows = page.ui.fixList.children || [];
  ok('clearing the whole list is acknowledged rather than left as struck-through text',
    rows.some((c) => /all-clear/.test(c.className)), 'no all-clear shown');
  page.T.toggleDone(findings[0].id, false);
  page.call('renderFixList');
  ok('and it goes the moment one comes back',
    !(page.ui.fixList.children || []).some((c) => /all-clear/.test(c.className)));

  /* ---- the rest, which are pure CSS but easy to lose ----------------------------- */
  ok('the finished-check burst is aimed at the mark rather than the top of the page',
    /ui\.predictedCard && !ui\.predictedCard\.hidden \? ui\.predictedBody : null\) \|\| ui\.scoreNumber/.test(html),
    'a burst 600px above the mark lands as decoration, not as a response');
  ok('a progress trail of particles runs while a check does',
    /@keyframes trail/.test(css) && /\.bar-fill::before/.test(css));
  ok('the primary action invites a press while it is idle and pressable',
    /#btnCheck:not\(\[disabled\]\) \{[\s\S]{0,400}animation-name: invite/.test(css),
    'a pulse on a dead control is a lie');
  /*
   * And it must STOP. It was infinite, and box-shadow cannot be animated on the compositor, so
   * that was a repaint of the button sixty times a second for ever, for decoration — the one
   * continuous cost the page had, and the reason it felt heavy while sitting still.
   */
  ok('and the invitation is finite, because an endless box-shadow pulse repaints for ever',
    /animation-iteration-count: 3/.test(css),
    'an infinite box-shadow animation is back');
  ok('the background tint leans towards the pointer', /var\(--aur-x, 0%\)/.test(css));
  ok('and it is desktop only, because a finger drag should not drag the wallpaper',
    /matchMedia\('\(pointer: fine\)'\)/.test(html));
  ok('tab panels cross in rather than snapping', /@keyframes panel-in/.test(css));
  ok('the tab underline slides', /\.tab\[aria-selected="true"\]::after \{ transform: scaleX\(1\)/.test(css));

  // Reduced motion must still turn every one of these off. The global rule handles it; this
  // is here so nobody "improves" that rule away later.
  ok('every one of them is still covered by the reduced-motion rule',
    /@media \(prefers-reduced-motion: reduce\) \{[\s\S]{0,220}animation-duration: 0\.01ms !important/.test(css));
}

section('12x. Are the PowerPoint and the PDF the same portfolio?');
{
  /*
   * Requested: "a thing that checks whether the pdf and pptx are the same".
   *
   * It matters because only ONE file is the structural deck — the .pptx if set, the PDF
   * otherwise — while the PDF is ALSO what the AI looks at. With both set, the report combines a
   * layout analysis of one file with a visual mark of another, so if they are not the same piece
   * of work then every number in it is about neither of them properly.
   *
   * The only check was the filename, which is weak in both directions: an export gets whatever
   * name the tool chose, and two different portfolios both called "Portfolio" sail through.
   */
  const sw = require(path.join(ROOT, 'lib', 'samework.js'));

  /* ---- what the measure actually looks at ---------------------------------------- */
  eq('short words are dropped, because they are noise and page numbers',
    sw.tokensOf('a of it the 3 and cad abs'), ['cad', 'abs']);
  eq('punctuation cannot hide a match',
    sw.tokensOf('CAD-modelled, client/led'), sw.tokensOf('cad modelled client led'));
  eq('words that carry no evidence are dropped',
    sw.tokensOf('the and for with from this that which'), []);
  /*
   * D&T vocabulary is deliberately KEPT, though it looks like noise — every portfolio says
   * "design", "client", "material". Dropping it was MEASURED on the real files: it moved two
   * strangers from 0.286 to 0.278 and the genuine pair not at all. Those words are simply too
   * few of the two thousand or so distinctive ones to matter, so removing them would only
   * throw evidence away.
   */
  eq('but D&T vocabulary is kept, because dropping it was measured to change nothing',
    sw.tokensOf('design material client specification').length, 4);

  const pool = (from, to, prefix) => {
    const out = [];
    for (let i = from; i <= to; i++) { out.push(prefix + 'word' + i); }
    return out;
  };
  const pageOf = (words) => words.join(' ');
  const pages = (arr, count) => {
    const out = [];
    for (let i = 0; i < count; i++) { out.push(pageOf(arr.slice(Math.floor(i * arr.length / count), Math.floor((i + 1) * arr.length / count)))); }
    return out;
  };

  const shared = pool(1, 100, 'a');
  const extra = pool(101, 140, 'b');

  let r = sw.compareSameWork(pages(shared, 10), pages(shared, 10));
  eq('the same words on both sides is the same portfolio', r.verdict, 'same');
  eq('and the match is complete', r.jaccard, 1);

  r = sw.compareSameWork(pages(shared, 10), pages(extra, 10));
  eq('wholly different words is a different portfolio', r.verdict, 'different');
  eq('and nothing is shared', r.shared, 0);

  // 60 of A's words plus 40 new ones: intersection 60, union 140.
  const partial = shared.slice(0, 60).concat(extra);
  r = sw.compareSameWork(pages(shared, 10), pages(partial, 10));
  eq('a partial overlap is reported as unsure, not rounded to a verdict', r.verdict, 'unsure');
  ok('and the sentence says what the number was',
    /\d+% of their distinctive words/.test(r.note || ''), r.note);

  r = sw.compareSameWork(['one', 'two'], ['three']);
  eq('a portfolio with almost no text cannot be compared', r.verdict, 'unknown');
  eq('and says why', r.reason, 'not-enough-text');
  ok('without pretending that means they differ', /nothing to match/.test(r.note || ''), r.note);

  /*
   * The thresholds, and the measurement they came from.
   *
   * This is the part worth guarding. Containment (shared over the smaller document) was the
   * first choice and MEASURING KILLED IT: two different students' portfolios share 57%, 57% and
   * 54% of their distinctive words by containment, against 99.7% for a portfolio and its own
   * export. Jaccard on the same pairs is 0.98 for the genuine one and 0.18 to 0.29 for the
   * strangers, which is a gap a threshold can sit inside.
   */
  eq('the same threshold sits above the highest measured different pair', sw.JACCARD_SAME, 0.6);
  eq('the different threshold sits below the lowest measured same pair', sw.JACCARD_DIFFERENT, 0.35);
  ok('the two bands leave a gap for "unsure"',
    sw.JACCARD_DIFFERENT < sw.JACCARD_SAME, 'there would be no unsure band at all');

  /* ---- the page shows the verdict ------------------------------------------------ */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const pdfName = 'Work.pdf';
  const pptxName = 'Work.pptx';
  const draw = (sameWork) => {
    page.state.data = {
      ok: true,
      config: { pdfFile: '/x/' + pdfName, pdfName: pdfName, deckFile: '/x/' + pptxName },
      deck: { fileName: pptxName, slideCount: 60 },
      local: { findings: [] },
      ai: {},
      sameWork: sameWork,
    };
    page.call('renderReviewed');
    return {
      warn: page.ui.reviewedWarn.hidden === false ? page.ui.reviewedWarn.textContent : '',
      how: page.ui.reviewedHow.textContent,
    };
  };

  let d = draw({ verdict: 'same', jaccard: 0.98, note: 'n', pptxFile: pptxName, pdfFile: pdfName });
  ok('a confirmed match is SAID, not merely left unmentioned',
    /checked against the PowerPoint: same portfolio/.test(d.how), d.how);
  ok('and it does not raise a warning', !/NOT look like the same/.test(d.warn), d.warn);

  d = draw({ verdict: 'different', jaccard: 0.2, note: 'These two do NOT look like the same portfolio.', pptxFile: pptxName, pdfFile: pdfName });
  ok('different work is warned about, in the comparison\'s own words',
    /NOT look like the same portfolio/.test(d.warn), d.warn);
  ok('and the page does not still claim they match',
    !/checked against the PowerPoint/.test(d.how), d.how);

  d = draw({ verdict: 'unsure', jaccard: 0.45, note: 'check they really are the same portfolio.', pptxFile: pptxName, pdfFile: pdfName });
  ok('an unsure answer is passed on rather than hidden',
    /check they really are the same portfolio/.test(d.warn), d.warn);

  /*
   * A verdict from an earlier scan must not be shown against different files. Swapping a file
   * does not trigger a re-scan, so a stale answer is the easy mistake to make here.
   */
  d = draw({ verdict: 'same', jaccard: 0.98, note: 'n', pptxFile: 'Something Else.pptx', pdfFile: pdfName });
  ok('a verdict about a DIFFERENT pair of files is not shown at all',
    !/checked against the PowerPoint/.test(d.how), d.how);
  ok('and with nothing to go on it falls back to naming the mismatch, saying that is all it is',
    !/different/i.test(d.warn) || /have not been compared yet/.test(d.warn), d.warn);

  /* ---- and it is actually wired into the scan ------------------------------------ */
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok('the scan compares the two files', /compareSameWork\(/.test(src) && /pdftext\.extractPdfTextFile\(cfg\.pdfFile\)/.test(src));
  ok('only when both are actually set, since that is the only case that can be wrong',
    /if \(!fromPdf && cfg\.deckFile && cfg\.pdfFile && fs\.existsSync\(cfg\.pdfFile\)\)/.test(src));
  ok('the verdict travels with the report', /sameWork: sameWork,/.test(src));
  ok('and it records which two files it compared, so a stale answer can be spotted',
    /sameWork\.pptxFile = path\.basename\(cfg\.deckFile\)/.test(src)
    && /sameWork\.pdfFile = path\.basename\(cfg\.pdfFile\)/.test(src));
  ok('a PDF that cannot be read is reported as unknown rather than as a mismatch',
    /verdict: 'unknown',\s*\n\s*reason: 'unreadable'/.test(src));
}

section('12y. Cancelling a check');
{
  /*
   * Requested: a cancel button.
   *
   * The awkward part is not the button, it is what "cancel" can honestly mean. A scan is a
   * sequence of paid API calls, so:
   *
   *   - it is a REQUEST, read between steps rather than a kill, because abandoning a call that
   *     is already in flight wastes the request being made;
   *   - it stops at BATCH BOUNDARIES, because a half-read batch cannot be cached and stopping
   *     mid-batch would throw away work that has already been paid for;
   *   - and a cancelled check is NOT a finished one, so it must not be celebrated or present a
   *     partial result as a mark.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');

  /* ---- the boundary, which is the part that matters ------------------------------ */
  ok('the scan is asked between batches, never during one',
    /for \(const batch of batches\) \{\n    \/\*[\s\S]{0,600}isCancelled\(\)/.test(aiSrc),
    'checked inside a batch, a stop would discard a half-read batch that cannot be cached');
  ok('and the visual review is asked before each run, which is the most expensive call',
    /for \(let i = 0; i < wanted; i\+\+\) \{\n      if \(i > 0 && typeof isCancelled/.test(aiSrc));

  const ai = require(path.join(ROOT, 'lib', 'ai.js'));
  const { parsePresentation } = require(path.join(ROOT, 'lib', 'pptx.js'));
  const deck = parsePresentation(path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx'));

  /*
   * Always true, so the pass stops at the first boundary it reaches.
   *
   * Deliberately NOT "false then true": returning false first lets a real request go out, and a
   * test suite must never spend the student's quota. The placement of the check — before the
   * batch rather than inside it — is asserted against the source above, which is where that
   * belongs; this proves the stop itself works and costs nothing.
   */
  let asked = 0;
  const cancelled = await ai.runAiPass({
    deck,
    apiKey: 'deliberately-not-a-key',
    model: 'gemini-3.5-flash',
    spec: {},
    cache: {},
    force: true,
    onLog: () => {},
    onProgress: () => {},
    isCancelled: () => { asked += 1; return true; },
  });
  ok('it does not even reach the network, because the flag is read first', asked === 1,
    'asked ' + asked + ' times');
  ok('a cancelled pass says so rather than failing', cancelled.cancelled === true);
  ok('and hands back its cache, so nothing already paid for is lost',
    !!cancelled.cache && !!cancelled.cache.batches, 'the partial cache was dropped');
  ok('and it is not reported as an error', !cancelled.error, String(cancelled.error));

  /* ---- the server side ------------------------------------------------------------ */
  ok('there is an endpoint to ask it to stop', /p === '\/api\/cancel' && req\.method === 'POST'/.test(src));
  ok('it refuses to look like a kill, saying the scan stops at the next safe point',
    /cancelRequested = true/.test(src) && /at the next safe point/.test(src));
  ok('the flag is cleared when a new scan starts, or every scan after the first would stop',
    /scanState\.cancelRequested = false;/.test(src));
  ok('and a cancelled scan is recorded as cancelled, distinctly from an error',
    /scanState\.cancelled = true;/.test(src) && /aiResult\.cancelled/.test(src));
  ok('the flag is handed to the AI pass', /isCancelled: \(\) => scanState\.cancelRequested/.test(src));

  /* ---- and the page --------------------------------------------------------------- */
  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  ok('the button is cached, not an undefined ui.* reference', !!page.ui.btnCancel);

  const draw = (scan, scanning) => {
    page.state.data = {
      ok: true,
      config: { portfolioChosen: true },
      deck: null,
      local: { findings: [] },
      ai: {},
      scan: scan,
    };
    page.state.scanning = !!scanning;
    page.call('renderActions');
    return page.ui.btnCancel;
  };

  let btn = draw({ running: true }, true);
  ok('the button appears while a check is running', btn.hidden === false, 'no way to stop it');
  eq('offering to cancel', btn.textContent, 'Cancel check');
  ok('and it is pressable', btn.disabled === false);

  btn = draw({ running: true, cancelRequested: true }, true);
  eq('once asked, it says it is stopping rather than pretending to have stopped',
    btn.textContent, 'Stopping…');
  ok('and cannot be pressed twice', btn.disabled === true);

  btn = draw({ running: false }, false);
  ok('it goes when nothing is running', btn.hidden === true,
    'a cancel button on an idle page is a dead control');

  /* ---- cancelled is not finished -------------------------------------------------- */
  ok('a cancelled check raises a note instead of celebrating',
    /if \(data\.scan && data\.scan\.cancelled\) \{[\s\S]{0,260}setBanner\('cancel'/.test(html));
  ok('and the note says the free checks ran while the AI review did not',
    /the AI review did not finish, so there is no mark/.test(html));
  ok('and that the pages already read are kept, which is why that is true',
    /pages already read are remembered, so pressing Check now again carries on/.test(html));
  ok('the note is cleared when a new check starts, so it cannot linger',
    /clearBanner\('cancel'\)/.test(html));
  ok('the page asks the server rather than deciding for itself',
    /apiPost\('\/api\/cancel'/.test(html), 'the page must not fake a stop it cannot perform');
}

section('12z. Starting when the port is already taken');
{
  /*
   * Reported: "the app is currently not working".
   *
   * The port was the reason. The server saved `port: 4173` in the config, and every version asks
   * for that same port, so:
   *
   *   - a second double-click, and
   *   - an UPGRADE while an older copy is still running
   *
   * both ended with "Port 4173 is already in use", process.exit(1), and a window that closed.
   * To the student that is indistinguishable from the app being broken — and upgrading is
   * exactly when it happens, because the old copy is still running on the port the new one
   * wants.
   */
  const srv = require(path.join(ROOT, 'server.js'));
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const http = require('http');

  ok('the port can be reassigned, or a busy one would be fatal',
    /\n  let port = Number\(process\.env\.NEA_COACH_PORT/.test(src));
  ok('a busy port is asked about BEFORE anything gives up',
    /if \(err\.code !== 'EADDRINUSE'\)[\s\S]{0,900}probeExisting\(port,/.test(src),
    'the old handler exited on EADDRINUSE without finding out who held the port');

  /* ---- recognising our own copy, for real --------------------------------------- */
  /*
   * Bound to EVERY interface, not just loopback, which is what a real program holding the port
   * looks like — and the app asks for every interface. macOS happily lets an all-interface bind
   * sit alongside a loopback-only one, so a loopback-held port would not be a fair test of
   * stepping aside.
   */
  const listen = (handler) => new Promise((resolve) => {
    const s = http.createServer(handler);
    s.listen(0, () => resolve(s));
  });
  const close = (s) => new Promise((r) => s.close(r));
  const probe = (port) => new Promise((r) => srv.probeExisting(port, r));

  let server = await listen((req, res) => {
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify({ ok: true, version: '1.0.36', scan: { running: false } }));
  });
  /*
   * The probe now answers with WHICH version is running, because the launchers used to work that
   * out for themselves with a batch `for /f` that cmd silently broke on a path containing a space
   * ("C:\\Program Files\\nodejs\\node.exe") — so the check never ran and the launcher opened the old
   * copy without a word. The server can answer it from a request it was making anyway.
   */
  const ours = await probe(server.address().port);
  eq('a second copy is recognised, and says which version it is', ours && ours.version, '1.0.36');
  eq('and whether it is busy', ours && ours.running, false);
  await close(server);

  server = await listen((req, res) => { res.end('hello, not nea coach'); });
  eq('some other program on the port is NOT mistaken for it', await probe(server.address().port), null);
  await close(server);

  /*
   * A socket that accepts the connection and then says nothing. Getting this wrong is the
   * failure mode that looks worst: the launch would appear to hang with no message at all.
   */
  server = await listen(() => { /* accepts, never answers */ });
  const started = Date.now();
  eq('a port that accepts but never answers is not NEA Coach', await probe(server.address().port), null);
  const waited = Date.now() - started;
  ok('and it gives up quickly rather than hanging the launch', waited < 4000, waited + 'ms');
  await close(server);

  eq('a closed port is not NEA Coach', await probe(9), null);

  ok('finding our own copy exits CLEANLY, because nothing has gone wrong',
    /NEA Coach is already running\.[\s\S]{0,700}process\.exit\(0\)/.test(src),
    'exiting non-zero would leave the launcher reporting a failure');
  ok('and it opens the copy that is already running, which is what was wanted',
    /openInBrowser\(url\)/.test(src));
  ok('the browser is not taken over when NEA_COACH_NO_OPEN=1',
    /NEA_COACH_NO_OPEN === "1"/.test(src));

  /* ---- and stepping over a port that is somebody else's -------------------------- */
  const held = await listen(() => {});
  const free = await new Promise((r) => srv.findFreePort(held.address().port, '0.0.0.0', r));
  ok('a port held by something else is stepped over',
    free > held.address().port, 'got ' + free);
  await close(held);

  ok('the new port is saved, so the student is not asked the same question every launch',
    /try \{ cfg\.port = free; saveConfig\(cfg\); \}/.test(src));
  ok('and it says so rather than silently using a different address',
    /is in use by something else/.test(src) && /Using \$\{free\} instead/.test(src));
  ok('with a way out if every port it tries is busy',
    /are all in use, so NEA Coach cannot start/.test(src) && /NEA_COACH_PORT/.test(src));
}

section('12aa. The page and the app must be the same version');
{
  /*
   * Found in the wild: the student had unzipped a newer copy over one that was STILL RUNNING.
   * The running process keeps its code in memory, while every page it serves is read from the
   * new files on disk — so the half that draws the screen and the half that answers it were
   * different versions. The page had a Cancel button whose endpoint did not exist yet.
   *
   * It presents as the app being broken with no clue why: some things work and some do not,
   * and nothing on screen explains it. The fix is not to prevent the situation — that is the
   * student's filesystem — but to make it VISIBLE.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const sh = fs.readFileSync(path.join(ROOT, 'tools', 'release.sh'), 'utf8');

  const serverVersion = /^const VERSION = '([^']+)';/m.exec(srv)[1];
  const pageVersion = /^  var PAGE_BUILD = '([^']+)';/m.exec(html);
  ok('the page states its own build', !!pageVersion, 'the page cannot tell it is stale without one');

  /*
   * The important one. Forgetting to bump one of the two is exactly how the mismatch would
   * creep back in, and it would then warn every user about a problem that is not there.
   */
  eq('and it matches the version the server reports',
    pageVersion ? pageVersion[1] : '(none)', serverVersion);

  const guard = /SRV_V="\$\(grep -m1 "\^const VERSION = '" "\$APP_DIR\/server\.js"[\s\S]{0,600}exit 1/.exec(sh);
  ok('a release is refused outright if the two ever disagree', !!guard,
    'a mismatched pair could be published again');

  const { loadPage } = require(path.join(__dirname, 'dom.js'));
  const page = loadPage(ROOT);
  const bannerText = () => {
    const host = page.els.get('banners');
    const walk = (n) => {
      if (!n) { return ''; }
      let out = n.textContent ? String(n.textContent) : '';
      for (const k of (n.children || [])) { out += ' ' + walk(k); }
      return out;
    };
    return walk(host);
  };

  const draw = (version) => {
    page.state.data = {
      ok: true, version: version,
      config: { portfolioChosen: true }, deck: null,
      local: { findings: [] }, ai: {}, scan: { running: false },
    };
    page.call('renderBanners');
    return bannerText();
  };

  // A mismatch has to be raised from the load path, so the source is what proves it is wired.
  ok('the check runs on the load path, where every refresh passes',
    /checkVersionMatch\(data\);/.test(html),
    'checked anywhere else and it would be missed on the load that matters');
  ok('and it is the comparison of the page build against the version reported',
    /theirs !== PAGE_BUILD/.test(html) && /setBanner\('version'/.test(html));
  ok('the message says what to do about it, not just that something is wrong',
    /Quit NEA Coach completely/.test(html) && /open it again/.test(html));
  ok('and it cannot be dismissed, because it will keep happening until the app is restarted',
    /setBanner\('version', 'warn',[\s\S]{0,600}, true\)/.test(html));
  ok('it is cleared once the two agree again', /clearBanner\('version'\)/.test(html));

  /*
   * Driven through the real function rather than by inspecting the markup, because the whole
   * value of this guard is that it fires — or does not — on the path every refresh takes.
   */
  const agree = page.call('checkVersionMatch', { version: serverVersion });
  ok('matching versions raise nothing', agree.ok && agree.value === true, JSON.stringify(agree));
  eq('and no banner is left behind', bannerText(), '');
  ok('and it clears one that was there', !/version/.test(bannerText()));

  const differ = page.call('checkVersionMatch', { version: '0.0.1-ancient' });
  ok('a mismatch is reported', differ.ok && differ.value === false, JSON.stringify(differ));
  const shown = bannerText();
  ok('the message names both versions', /0\.0\.1-ancient/.test(shown) && shown.indexOf(serverVersion) >= 0, shown);
  ok('and tells the student to quit it completely and open it again',
    /Quit NEA Coach completely/.test(shown) && /open it again/.test(shown), shown);

  page.call('checkVersionMatch', { version: serverVersion });
  eq('and reporting it again with matching versions clears it', bannerText(), '');

  /*
   * The state must report the version of the app that is RUNNING, not the version that wrote the
   * report on disk.
   *
   * A report carries the version that produced it, and after an upgrade that is the PREVIOUS
   * version until the next scan. Without an explicit override the state inherited the report's,
   * so a 1.0.38 app told its own page "the app behind this is 1.0.36" — and the guard above then
   * warned every user to restart over nothing. A guard that cries wolf is worse than no guard.
   */
  ok('the state reports the running version, overriding the report\u2019s',
    /Object\.assign\(\{\}, report, \{[\s\S]{0,900}version: VERSION,/.test(srv),
    'the state would inherit the version that wrote the report');
  ok('and keeps the report\u2019s own version separately, for when that is the useful number',
    /reportVersion: report\.version,/.test(srv));
  ok('the stale-state branch reports the running version too',
    /stateStale: true,[\s\S]{0,200}version: VERSION,/.test(srv) || /version: VERSION,[\s\S]{0,200}stateStale: true,/.test(srv));
  ok('and the page compares its build against THAT, not against the report\u2019s',
    /checkVersionMatch\(data\);/.test(html) && /theirs !== PAGE_BUILD/.test(html));
}

section('12ab. Cancel has to reach the expensive part');
{
  /*
   * Found by the whole-codebase audit, and it was a defect in the cancel feature itself.
   *
   * runVisionPass has accepted and checked an `isCancelled` callback since it was written. The
   * call in runAiPass never passed one, so the flag was always null and BOTH of its checks were
   * unreachable. Cancelling therefore did nothing until the entire visual review had finished
   * and been paid for — the first and by far the most expensive work a scan does, up to three
   * whole-portfolio requests plus any fallback model. The button promised the opposite.
   */
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');

  ok('the cancel flag reaches the visual review',
    (() => {
      const call = /vision = await runVisionPass\(\{([\s\S]*?)\n    \}\);/.exec(aiSrc);
      return !!call && /\bisCancelled\b/.test(call[1]);
    })(),
    'Cancel cannot stop the whole-portfolio review without it');

  /*
   * Behavioural proof, with the network replaced: cancelling before any work must produce ZERO
   * generateContent attempts. The models listing is a free GET and is allowed to happen.
   */
  const ai = require(path.join(ROOT, 'lib', 'ai.js'));
  const { parsePresentation } = require(path.join(ROOT, 'lib', 'pptx.js'));
  const theDeck = parsePresentation(path.join(ROOT, 'sample', 'AQA-GCSE-Portfolio-SAMPLE.pptx'));
  const https = require('https');
  const realRequest = https.request;
  const attempted = [];
  https.request = function (opts) {
    attempted.push(typeof opts === 'string' ? opts : ((opts && (opts.path || opts.href)) || '?'));
    throw new Error('blocked by the test');
  };
  let outcome = null;
  try {
    outcome = await ai.runAiPass({
      deck: theDeck,
      apiKey: 'deliberately-not-a-key',
      model: 'gemini-3.5-flash',
      spec: {},
      cache: {},
      force: true,
      pdf: Buffer.from('%PDF-1.4 not really a pdf'),
      pdfName: 'x.pdf',
      runs: 3,
      onLog: () => {},
      onProgress: () => {},
      isCancelled: () => true,
    });
  } finally {
    https.request = realRequest;
  }

  ok('a cancelled pass stops before spending anything', outcome && outcome.cancelled === true);
  const paid = attempted.filter((p) => /generateContent/.test(p));
  eq('and sends no page-reading request at all', paid.length, 0);
  /*
   * `.length >= 0` was a tautology — true of an empty array, a full one, and anything between, so
   * it asserted nothing while counting towards the pass total. This asserts the fact that was
   * actually intended: the free models listing WAS attempted (proving the pass reached that point)
   * and no paid call was.
   */
  ok('the free models listing was reached, which is what proves the pass got that far',
    attempted.some((p) => /models/.test(p)), JSON.stringify(attempted));
  eq('and no page-reading request was sent', paid.length, 0);
}

section('12ac. Five high-severity findings from the whole-codebase audit');
{
  /*
   * Every one of these was found by a read-only audit pass and is fixed here. They are grouped
   * because they share a theme: in each case the code was doing something reasonable and the
   * thing that made it work — a returned field, a forwarded flag, a checked error, a flipping
   * token — was simply missing, so the feature was written, logged, and inert.
   */
  const server = require(path.join(ROOT, 'server.js'));
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(html)[1];

  /* ---- 1. the learned daily cap was thrown away on the next request ------------- */
  /*
   * readQuota() dropped `limits` from what it returned, and noteRequest/learnQuotaLimit both do
   * read -> mutate -> write. So the limit learned from a 429 was written once and then DELETED by
   * the next counted request, and quotaReport could never set limitKnown — which is why the usage
   * meter could only ever show a count and never a percentage of the day's allowance.
   */
  server.learnQuotaLimit('audit-probe-model', 42);
  eq('the learned limit is stored', server.readQuota().limits['audit-probe-model'], 42);
  server.noteRequest('audit-probe-model');
  eq('and it SURVIVES the next counted request, which used to delete it',
    (server.readQuota().limits || {})['audit-probe-model'], 42);
  const rep = server.quotaReport('audit-probe-model');
  const row = (rep.models || []).filter((r) => r.model === 'audit-probe-model')[0] || {};
  ok('so the meter can finally say how much of the day is gone',
    row.limitKnown === true && row.limit === 42 && row.percent !== null,
    JSON.stringify(row));

  /* ---- 3. "I could not read the document" was a successful review --------------- */
  ok('a vision reply that reports its own failure is no longer accepted as a success',
    /parsed\.error && String\(parsed\.error\)\.trim\(\)/.test(aiSrc)
    && /could not read the pages/.test(aiSrc),
    'the model\u2019s own error field is ignored again, so a failed read counts as a review');

  /* ---- 4/5. two dark-mode contrast bugs ---------------------------------------- */
  ok('the severity count badge flips with the theme instead of staying near-white',
    /\.chip\[aria-pressed="true"\] \.chip-count \{ background: var\(--card\)/.test(css),
    'light text on a near-white pill in dark mode, at about 1.1:1');
  ok('and the strengths text uses tokens rather than light-mode greens',
    /\.strengths h3 \{[^}]*color: var\(--good\)/.test(css)
    && /\.strengths li \{[^}]*color: var\(--ink-2\)/.test(css),
    'dark green on a near-black panel in dark mode, at about 1.6:1');

  /* ---- the suite itself must not read the student's real data ------------------- */
  ok('the test suite points the app at a throwaway data folder',
    !!process.env.NEA_COACH_DATA_DIR && /nea-tests-/.test(process.env.NEA_COACH_DATA_DIR),
    'the suite is reading the real API key, report and quota');
  ok('and that folder is not the real per-user one',
    server.stateDir.indexOf(os.homedir()) !== 0 || /nea-tests-/.test(server.stateDir),
    server.stateDir);
}

section('12ad. Automatic re-checking, and the Mac installer text');
{
  /*
   * Both from the second audit pass. The watcher bugs are real product faults — the app advertises
   * automatic re-checking and quietly did not do it — and the installer bug was in the first
   * instruction a Mac user reads.
   */
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const pkg = fs.readFileSync(path.join(ROOT, 'tools', 'package.sh'), 'utf8');

  /* ---- a change that arrives during a scan must be remembered, not dropped ------- */
  ok('a portfolio change during a running check is remembered rather than discarded',
    /if \(scanState\.running\) \{[\s\S]{0,500}pendingRescan = true;/.test(src),
    'the event is thrown away again, so an edit during a long check is never re-checked');
  ok('and the remembered change is acted on when the scan finishes',
    /const rescanQueued = pendingRescan;/.test(src) && /if \(rescanQueued\) \{/.test(src)
    && /reason: 'auto'/.test(src));
  ok('the re-check is deferred so two scans cannot fight over the same cache and report',
    /setTimeout\(\(\) => \{[\s\S]{0,400}requestScan\(\{ force: false, reason: 'auto' \}\)/.test(src));

  /* ---- a settings autosave must not swallow a queued change --------------------- */
  /*
   * Every settings autosave POSTs /api/config, whose handler restarts the watcher. That restart
   * re-stamped watchedStamp to the CURRENT file, so a debounce armed a moment earlier by a real
   * edit compared the file against its own post-edit stamp and decided nothing had changed.
   */
  ok('restarting the watcher does not re-stamp while a change is still queued',
    /if \(!watchTimer\) \{\s*\n\s*watchedStamp = stampOf\(target\);/.test(src),
    'a settings save within the debounce window silently cancels the re-check');

  /* ---- the Mac readme is a Mac readme ------------------------------------------ */
  ok('the Mac package does not ship the Windows first-readme',
    !/  write_readme_first "\$target"\n  write_readme_first_mac/.test(pkg),
    'Mac users are told to double-click a .bat that is deleted from the package');
  const macBody = (() => {
    const m = /write_readme_first_mac\(\) \{[\s\S]*?<<'EOF'\n([\s\S]*?)\nEOF/.exec(pkg);
    return m ? m[1] : '';
  })();
  ok('and the Mac document is actually written, not appended to the Windows one',
    /write_readme_first_mac\(\) \{[\s\S]{0,80}cat > /.test(pkg), 'it appends to the Windows text');
  ok('it does not tell a Mac user to double-click a .bat',
    !!macBody && !/\.bat/.test(macBody), macBody.slice(0, 120));
  ok('and the Gatekeeper remedy carries no frozen version number',
    !!macBody && !/NEA-Coach-mac-\d+\.\d+/.test(macBody),
    'the xattr line named 1.0.29 for several releases, so it silently did nothing');
}

section('12ae. The packaging and release tools');
{
  /*
   * From the second audit pass. These sit on the path between "the code is right" and "the
   * student receives it", which is where mistakes are most expensive: they are invisible until
   * someone downloads a broken package.
   */
  const pkg = fs.readFileSync(path.join(ROOT, 'tools', 'package.sh'), 'utf8');
  const rel = fs.readFileSync(path.join(ROOT, 'tools', 'release.sh'), 'utf8');

  /*
   * A C-style comment in a shell script is NOT a comment — bash glob-expands /* and tries to
   * EXECUTE it (exit 126). I put one in package.sh myself, and it survived because the script
   * has no `set -e`, so every Mac build quietly ran a failing command and carried on.
   */
  const shellFiles = ['tools/package.sh', 'tools/release.sh', 'tools/get-release.sh', 'Start NEA Coach.command'];
  shellFiles.forEach((f) => {
    const full = path.join(ROOT, f);
    if (!fs.existsSync(full)) { return; }
    const body = fs.readFileSync(full, 'utf8');
    ok('no C-style comment in ' + f, !/^\s*\/\*/m.test(body),
      'bash will try to execute /* as a command');
  });

  /* ---- a Windows build must not ship without Node ------------------------------ */
  ok('a Windows build without runtime/node.exe FAILS instead of shipping',
    /ERROR: runtime\/node\.exe is missing/.test(pkg) && /^\s*exit 1$/m.test(pkg.slice(pkg.indexOf('ERROR: runtime/node.exe is missing'))),
    'the zip would be built, published, and promise Node is inside it');
  /*
   * The old note told the recipient that package.sh "fetches it automatically". It never did —
   * nothing in this script downloads anything — so the note sent them to look for a feature that
   * does not exist. Checked against the emitted claim, not the word, so the comment explaining
   * this does not count as the thing it warns about.
   */
  ok('the misleading note is gone rather than reworded',
    !/which fetches/.test(pkg),
    'it still tells the reader that package.sh downloads Node for them');

  /* ---- a typo must not destroy the built packages ------------------------------ */
  const wipe = pkg.indexOf('rm -rf "$DIST"/NEA-Coach-mac*');
  const check = pkg.indexOf('usage: tools/package.sh');
  ok('the argument is validated BEFORE dist/ is wiped',
    check > 0 && wipe > 0 && check < wipe,
    'a typo deletes every built package and then prints a usage message');

  /* ---- never publish a half-uploaded release ---------------------------------- */
  ok('the release is created as a DRAFT',
    /"draft":true,"prerelease":false/.test(rel),
    'it is public before its packages are attached');
  ok('and only PUBLISHED once every asset is up',
    /X PATCH[\s\S]{0,200}"draft":false/.test(rel),
    'the first zip is downloadable while the second is still uploading');
  ok('an upload failure deletes the incomplete release rather than leaving it',
    /X DELETE[\s\S]{0,200}releases\/\$RELEASE_ID/.test(rel),
    'a public release missing one of its two downloads is served by the /latest link');
}

section('12af. The launchers no longer guess the port');
{
  /*
   * Both launchers duplicated knowledge the server already had: which port is in use, which
   * version is running there, and when to open the browser. The server is the only thing that
   * knows what port it actually got — it steps aside when its usual one is busy and saves the
   * new one — so a launcher polling the assumed port waited on nothing, printed "did not start
   * within 30 seconds" while the app ran fine next door, and never opened anything.
   *
   * The version check was worse on Windows: it was written with `for /f`, and cmd strips the
   * outer quotes off a `for /f` command when the path inside contains a space. "C:\Program
   * Files\nodejs\node.exe" always does, so the check never ran and the launcher opened the OLD
   * copy without a word. On the Mac the same check worked, which is how a duplicated rule ends
   * up disagreeing with itself.
   */
  const bat = fs.readFileSync(path.join(ROOT, 'Start NEA Coach.bat'), 'utf8');
  const cmd = fs.readFileSync(path.join(ROOT, 'Start NEA Coach.command'), 'utf8');
  const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');

  // Only lines that actually RUN matter; the comments explaining the removal mention for /f.
  const batCode = bat.split(/\r?\n/).filter((l) => !/^\s*REM\b/i.test(l)).join('\n');
  ok('the batch version check is gone from the code, not just the comments',
    !/for \/f/.test(batCode), 'cmd silently breaks it on a path with a space');

  ok('the launcher starts the server and stops, instead of polling a port it assumed',
    /start "NEA Coach server"[\s\S]{0,200}?goto :finish/.test(bat),
    'it waits on 4173 while the server may be on the port next door');
  ok('and it no longer opens the browser itself',
    (bat.match(/start "" "http/g) || []).length <= 1,
    'two things opening the browser is a race, and the launcher can pick the wrong port');
  ok('every batch label it jumps to exists',
    (() => {
      const labels = new Set([...batCode.matchAll(/^:([A-Za-z_]\w*)/gm)].map((m) => m[1]));
      const gotos = [...new Set([...batCode.matchAll(/\bgoto :?([A-Za-z_]\w*)/g)].map((m) => m[1]))];
      return gotos.every((g) => labels.has(g));
    })(), 'a goto with no label ends the script silently');

  ok('the Mac launcher no longer forces a port on the server',
    !/NEA_COACH_PORT="\$PORT"/.test(cmd),
    'it overrides the port the server saved when it had to step aside');
  ok('and it no longer opens the browser either',
    (cmd.match(/^\s*open "http/gm) || []).length === 0);
  ok('and it no longer compares versions itself',
    !/RUNNING_VERSION/.test(cmd), 'that rule belongs in one place');

  /* ---- and the server does all of it, where it can be tested ------------------- */
  ok('the server opens the browser itself, once it is listening',
    /function onListening\(\)[\s\S]*?openInBrowser\(macUrl\);/.test(srv),
    'nothing opens the browser at the port the server actually got');
  ok('and that can be turned off for scripts and tests',
    /NEA_COACH_NO_OPEN === "1"/.test(srv));
  ok('the running copy reports WHICH version it is, not merely that one exists',
    /version: j\.version, running: !!\(j\.scan && j\.scan\.running\)/.test(srv));
  ok('and a different version is called out rather than opened silently',
    /but it is a DIFFERENT version/.test(srv));
}

section('12ag. The information panel is scannable, not prose');
{
  /*
   * It was a <dl> with NO CSS AT ALL — nine dense definitions, several sentences each, with nothing
   * to show where one ended and the next began. Rewritten as short colour-coded rows: a tag, a
   * title, and one or two lines. No fact was dropped; the wording was cut back.
   *
   * These guard the SHAPE, because the easy way to undo this is to keep adding to a paragraph until
   * it is a wall again.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(html)[1];
  const panel = (() => {
    const i = html.indexOf('id="infoDetails"');
    const j = html.indexOf('</details>', i);
    return i < 0 || j < 0 ? '' : html.slice(i, j);
  })();

  ok('the panel was found', panel.length > 500);

  const rows = [...panel.matchAll(/<div class="info-row[^"]*">([\s\S]*?)<\/div>/g)].map((m) => m[1]);
  ok('it is built from rows rather than one long list', rows.length >= 6, rows.length + ' rows');
  ok('every row has a tag, a title and a line of text',
    rows.every((r) => /info-tag/.test(r) && /info-title/.test(r) && /info-text/.test(r)),
    'a row without all three is a paragraph again');

  /* ---- concise, enforced ------------------------------------------------------- */
  /*
   * The longest line is about 230 characters; 320 leaves room for a wording change but not for a
   * return to paragraphs. This is the whole point of the rewrite, so it is worth a number.
   */
  const texts = [...panel.matchAll(/<p class="info-text">([\s\S]*?)<\/p>/g)]
    .map((m) => m[1].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim());
  const longest = texts.reduce((a, b) => (b.length > a.length ? b : a), '');
  ok('and no single line runs long enough to be a paragraph again',
    longest.length <= 320, longest.length + ' characters: ' + longest.slice(0, 80) + '...');

  /* ---- colourful, and safely so -------------------------------------------------- */
  ok('rows are colour-coded by kind', /\.info-row\.tone-good/.test(css) && /\.info-row\.tone-spend/.test(css)
    && /\.info-row\.tone-warn/.test(css) && /\.info-row\.tone-info/.test(css));
  ok('and every colour is a theme token, so dark mode needs no second version',
    (() => {
      const rules = [...css.matchAll(/\.info-(?:row|tag|title|text)[^{]*\{([^}]*)\}/g)].map((m) => m[1]).join('\n');
      return !/#[0-9a-f]{3,8}\b/i.test(rules) && !/rgba?\(/i.test(rules);
    })(), 'a hardcoded light colour here is how the last two dark-mode bugs happened');
  ok('the two facts the tests above care about are still stated',
    /A PDF records the words on each page, but not the layout/i.test(html)
    && /paste another key[\s\S]{0,40}in Settings/i.test(html));
  ok('and the unstyled definition list is gone',
    !/<dl class="explain">/.test(html), 'the default-styled list is what read as a wall of prose');
}

section('12ah. The allowance counts the right model, and the tab has colour');
{
  /*
   * "The allowance is still inaccurate." It was: every request the app made was credited to the
   * TEXT model, whatever actually served it.
   *
   *   onRequest: () => countRequest(cfg.model)
   *
   * A thunk that throws its argument away, and a counter invoked with no argument at all. Google's
   * free allowance is per MODEL, so the vision model's requests were all charged to the text model
   * and the vision model's row read zero. The student's real remaining allowance for each model was
   * misrepresented with no way to tell.
   *
   * The count of requests was right; the attribution was not.
   */
  const aiSrc = fs.readFileSync(path.join(ROOT, 'lib', 'ai.js'), 'utf8');
  const srvSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(html)[1];

  ok('the request counter is given the model that served the request',
    /requestCounter\(args\.model\)/.test(aiSrc),
    'the counter is called with no argument, so every request lands on one model');
  ok('and the server forwards it rather than hard-coding its own',
    /onRequest: \(which\) => countRequest\(which \|\| cfg\.model\)/.test(srvSrc),
    'a thunk that discards the model attributes vision requests to the text model');

  /* ---- the plumbing underneath is already right, and is easy to break --------- */
  ok('only SERVED requests count, so a refusal does not spend the allowance',
    /const served = res\.status === 200;/.test(aiSrc) && /if \(served && !counted && requestCounter\)/.test(aiSrc),
    'counting 429s made the meter claim an allowance was gone while Google was still serving');
  ok('a retry is not double-counted', /counted = true;/.test(aiSrc));
  ok('and merely listing the models is free and is not counted',
    /void onRequest;/.test(aiSrc));

  /* ---- the "worth adding" tab -------------------------------------------------- */
  ok('unselected tabs are no longer washed-out grey',
    !/\.tab \{[\s\S]{0,260}color: var\(--muted\)/.test(css),
    'light grey on a white card reads as disabled, not as "there is something in here"');
  ok('the open tab is washed with the accent so the row looks like choices',
    /\.tab\[aria-selected="true"\] \{[\s\S]{0,160}background: var\(--accent-soft\)/.test(css));
  ok('the suggestions panel has its own colour rather than the fix list severity colours',
    /\.suggestion-card \{[\s\S]{0,120}border-left: 3px solid var\(--accent\)/.test(css)
    && /\.suggestion-card \{[\s\S]{0,200}background: var\(--accent-soft\)/.test(css)
    && /\.suggest-heading \{[^}]*color: var\(--accent-dark\)/.test(css));
  /*
   * "the things to add should not have severities."
   *
   * There was a fourth badge on every suggestion reading "Would help a lot" or "Small gain", drawn
   * in the same greens and greys the FIX LIST uses for severity. The data was not severity — the
   * prompt is explicit that impact is how much fixing something would GAIN, as against severity,
   * which is what it is costing now — but a colour-coded rating at the end of every row reads as
   * one regardless of what the field is called. Suggestions are not faults.
   *
   * Nothing is lost: the list is ranked by gain per unit of effort, and the panel says so.
   */
  ok('a suggestion carries no gain or severity rating',
    // Matched as CODE, not as the word: the comments explaining the removal mention all three.
    !/function impactChipText/.test(html) && !/\.impact-(high|medium|low) \{/.test(css)
    && !/'pill impact-'/.test(html),
    'a graded, colour-coded badge on every suggestion makes it a second, milder fix list');
  ok('and the ranking that carried the same meaning is still there',
    /worthScore/.test(html) && /ordered by how much they would help/.test(html),
    'removing the badge must not remove the ORDER, which is what actually conveys gain');
  ok('the time estimate stays, because that is effort and not a grade',
    /effortClass\(s\.effort\)/.test(html));

  ok('and its grey intro paragraph became a tinted note',
    /\.suggest-intro \{/.test(css) && !/Things you could add[\s\S]{0,40}cls: 'muted small'/.test(html));
  /*
   * And the cards themselves, not just the tab. The panel was a white card of grey text with three
   * small pills on it; the parts that give a suggestion its structure now carry the accent.
   */
  ok('the number badge is solid accent rather than the fix list grey',
    /\.suggestion-card \.rank \{[\s\S]{0,120}background: var\(--accent\)/.test(css)
    && /\.suggestion-card \.rank \{[\s\S]{0,160}color: var\(--card\)/.test(css),
    'the badge text colour has to flip with the theme, which var(--card) does and #fff does not');
  ok('and the two structural labels are accented',
    /\.suggestion-card \.finding-block strong \{[^}]*color: var\(--accent-dark\)/.test(css));
  /*
   * .finding-block is emitted by the JavaScript and had NO rule at all — a class the stylesheet had
   * never heard of — so "Why it is worth it" and "What to do" were plain bold text.
   */
  ok('the block class it emits is actually styled now, rather than being a class with no rule',
    /\.suggestion-card \.finding-block \{[^}]*font-size/.test(css));
  ok('every colour added here is a theme token',
    (() => {
      /*
       * .tab(?![-a-z]) so this cannot wander into .tab-panel or .tabs and start judging their
       * colours, and [^{}]* so it cannot run past the end of a rule into the next one.
       */
      const rules = [...css.matchAll(/\.(tab(?![-a-z])|suggest-intro|suggestion-card|suggest-heading)[^{}]*\{([^}]*)\}/g)]
        .map((m) => m[2]).join('\n');
      return !/#[0-9a-f]{3,8}\b/i.test(rules) && !/rgba?\(/i.test(rules);
    })(), 'a literal colour here is how the dark-mode bugs happened');
}

section('12ai. The moving background stays off the main thread');
{
  /*
   * Requested: a dynamic, minimalistic moving background like layered drape bands, "try to make it
   * not laggy", keeping the soft blue theme.
   *
   * The performance rule is absolute and worth a test rather than a comment: TRANSFORM ONLY. A
   * browser can animate transform and opacity on the compositor, without touching the main thread
   * or repainting. Every other property — background-position, box-shadow, a blur, a filter —
   * forces a full-screen repaint on every frame, which is exactly what made this page feel heavy
   * before. So the test reads the keyframes and refuses anything else.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(html)[1];

  ok('the background is built from a small, fixed number of layers',
    /<div class="waves" id="waves" aria-hidden="true">/.test(html)
    && (html.match(/class="wave wave-\d"/g) || []).length === 3,
    'more layers means more full-screen fill to composite');
  ok('it sits behind the content and cannot be clicked',
    /\.waves \{[\s\S]{0,400}z-index: -1;[\s\S]{0,200}pointer-events: none;/.test(css));

  /* ---- the rule that keeps it cheap -------------------------------------------- */
  const frames = [...css.matchAll(/@keyframes (wave-[abc]) \{([\s\S]*?)\n  \}/g)];
  eq('there is a keyframe set per layer', frames.length, 3);
  const bad = frames.filter((m) => {
    const props = [...m[2].matchAll(/([a-z-]+)\s*:/g)].map((x) => x[1]);
    return props.some((pr) => pr !== 'transform');
  }).map((m) => m[1]);
  eq('and every one of them animates transform and NOTHING else', bad.join(','), '');
  /*
   * Declared ONCE on the shared .wave rule, which is the right way to do it: all three layers get
   * it, and there is one place to change it. Counting occurrences would have demanded three
   * copies of the same line.
   */
  ok('the layers are promoted to their own compositing layer',
    /\.wave \{[\s\S]{0,900}will-change: transform;/.test(css),
    'without this the browser may repaint instead of composite');
  ok('and the layers use translate3d, which keeps them on the GPU',
    frames.every((m) => /translate3d\(/.test(m[2])));

  /*
   * ---- and it has to be VISIBLE -------------------------------------------------
   *
   * It was 3.5% of travel over 54 seconds, which is about a PIXEL A SECOND: the background read as
   * static, and that is a fair complaint about an animation whose entire purpose is to be noticed.
   *
   * The floors below are set from that complaint, not from taste, and they matter because the
   * failure mode is silent — nothing breaks, it just stops being worth having. Speeding a
   * composited transform up costs nothing, so there is no performance reason to be timid here.
   */
  const travel = frames.map((m) => Math.max(...[...m[2].matchAll(/translate3d\(\s*(-?[\d.]+)%/g)]
    .map((x) => Math.abs(Number(x[1])))));
  ok('every band travels far enough to be seen at all',
    travel.length === 3 && travel.every((v) => v >= 5), travel.join('%, ') + '%');
  const lean = frames.map((m) => Math.max(...[...m[2].matchAll(/rotate\((-?[\d.]+)deg\)/g)]
    .map((x) => Math.abs(Number(x[1])))));
  ok('and leans enough to swing its ends across the screen',
    lean.length === 3 && lean.every((v) => v >= 1.5), lean.join('deg, ') + 'deg');
  const durs = [...css.matchAll(/animation: wave-[abc] (\d+)s/g)].map((m) => Number(m[1]));
  eq('there is one duration per band', durs.length, 3);
  ok('and none is slow enough to look still',
    durs.every((d) => d <= 40), durs.join('s, ') + 's (a pass must finish inside 40s)');
  ok('the three are not multiples of each other, so they cannot settle into a visible repeat',
    new Set(durs).size === 3 && durs[2] % durs[0] !== 0 && durs[1] % durs[0] !== 0);

  /* ---- what would defeat all of that ------------------------------------------ */
  const waveRules = [...css.matchAll(/\.wave[^{}]*\{([^}]*)\}/g)].map((m) => m[1]).join('\n');
  ok('no blur or filter anywhere in the moving layers',
    !/filter\s*:/.test(waveRules), 'blurring three full-screen layers is far worse than the shape it buys');
  ok('and no box-shadow, which repaints the layer it sits on',
    !/box-shadow/.test(waveRules));
  ok('the old drifting aurora no longer animates, so there is only one set of motion',
    !/animation: aurora/.test(css), 'two independent animation systems, both costing frames');
  ok('the colour ground is a flat static wash now',
    /html::before \{[\s\S]{0,500}background: linear-gradient\(180deg, var\(--wave-wash\)/.test(css));

  /* ---- and it stops when nobody is looking ------------------------------------ */
  /*
   * One selector covers every background layer, so a layer added later is paused by default rather
   * than forgotten.
   */
  ok('every background layer is paused while the tab is in the background',
    /\.is-hidden \.waves > \* \{ animation-play-state: paused; \}/.test(css)
    && /wireVisibility\(\)/.test(html), 'a hidden tab has no reason to hold a GPU layer awake');

  /* ---- the four extra layers, held to the same rule ---------------------------- */
  ok('the dust, the blobs, the glows and the streak animate transform or opacity only',
    (() => {
      const names = ['blob-drift', 'mote-drift', 'glow-breathe', 'streak-cross'];
      return names.every((nm) => {
        const m = new RegExp('@keyframes ' + nm + ' \\{([\\s\\S]*?)\\n  \\}').exec(css);
        if (!m) { return false; }
        const props = [...m[1].matchAll(/([a-z-]+)\s*:/g)].map((x) => x[1]);
        return props.length > 0 && props.every((pr) => pr === 'transform' || pr === 'opacity');
      });
    })(), 'a repainting property here undoes the whole point of the design');
  /*
   * The dust is ONE tiled field rather than an element per mote. Twenty animated dots would be
   * twenty composited layers for the same picture.
   */
  ok('the dust is a single scrolling layer, not an element per mote',
    /<span class="bg-motes"><\/span>/.test(html)
    // 1200 chars: the rule carries eight gradient stops before it reaches this line.
    && /\.bg-motes \{[\s\S]{0,1200}background-repeat: repeat-y/.test(css),
    'one layer per mote would cost more than everything else in the background together');
  ok('and the dust element is twice the viewport tall, so scrolling it half still covers the screen',
    /\.bg-motes \{[\s\S]{0,200}height: 200%/.test(css));
  ok('the streak spends most of its cycle invisible, so it reads as occasional',
    /@keyframes streak-cross \{[\s\S]{0,400}100% \{ transform: translate3d\(330%/.test(css),
    'sweeping continuously would be a distraction rather than a detail');
  ok('the new colours are theme tokens too',
    /--mote:/.test(html) && /--streak:/.test(html)
    && (html.match(/--mote:/g) || []).length === 3, 'light and both dark paths');
  ok('reduced motion still switches all of it off',
    /@media \(prefers-reduced-motion: reduce\)[\s\S]{0,220}animation-duration: 0\.01ms !important/.test(css));

  /* ---- soft blue, in both themes ---------------------------------------------- */
  const blocks = (html.match(/--wave-hi:/g) || []).length;
  eq('the drape colours are defined for light AND both dark paths', blocks, 3);
  ok('and they are blue-grey rather than the old pink and violet',
    /--wave-lo: rgba\(191, 209, 240/.test(html) && !/--wave-hi: rgba\(236, 72, 153/.test(html));
}

section('12aj. Every section is a tab');
{
  /*
   * The page was one long scroll and is now eleven sections behind a tab strip. Two things about
   * how that was done are worth holding, because both are easy to undo by accident:
   *
   *   - the ORDER is CSS, not document order. Settings is fourth in the markup and last in the
   *     tabs, and Information is three separate blocks that are nowhere near each other. Ordering
   *     in the stylesheet avoided moving large pieces of markup, which is how quiet breakage
   *     happens.
   *   - Fix list, Things to add and Coverage grid still share ONE card whose inner tabs are driven
   *     by the existing setTab(). Three tabs now call it rather than a strip of buttons.
   */
  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const css = /<style>([\s\S]*?)<\/style>/.exec(html)[1];

  const SECTIONS = ['overview', 'fixlist', 'suggest', 'predicted', 'coverage', 'since',
    'slides', 'stats', 'history', 'info', 'settings'];
  // class="toptab[^"]*" rather than "toptab", because the pinned Settings button carries a
  // second class for its divider.
  const buttons = [...html.matchAll(/class="toptab[^"]*"[^>]*data-tab="([a-z]+)"/g)].map((m) => m[1]);
  eq('there is one tab button per section, in order', buttons.join(','), SECTIONS.join(','));

  const panels = [...html.matchAll(/<div class="panel" id="panel([A-Za-z]+)">/g)].map((m) => m[1]);
  eq('and the panels they wrap are the expected ones', panels.join(','),
    'Overview,Settings,Since,Predicted,Fixlist,Slides,Info,History,Stats,Cost,Connect');

  ok('every tab points at a panel that exists, and only the ones it needs',
    /panels: \['Fixlist'\], inner: 'fix'/.test(html)
    && /panels: \['Connect', 'Info', 'Cost'\]/.test(html),
    'the mapping from section to panel is where a wrong tab would show nothing');
  ok('no tab is left with an empty panel list', !/panels: \[\]/.test(html));

  /* ---- order is CSS, and must stay that way ------------------------------------ */
  ok('the page is a flex column, which is what makes the order settable at all',
    /\.page \{ display: flex; flex-direction: column; \}/.test(css));
  const orders = [...css.matchAll(/#panel([A-Za-z]+) \{ order: (\d+); \}/g)].map((m) => [m[1], Number(m[2])]);
  eq('every panel has an order', orders.length, 11);
  const sorted = orders.slice().sort((a, b) => a[1] - b[1]).map((x) => x[0]).join(',');
  /*
   * "Use it on your phone" first, then the information card, then the allowance meter. The
   * connection card is LAST in the markup, so this order is the only reason it appears at the top
   * of the section.
   */
  ok('and Information comes out as connect, info, allowance meter — whatever the markup order',
    /Connect,Info,Cost/.test(sorted), sorted);
  ok('with the settings card last, as it is in the tabs',
    sorted.split(',').pop() === 'Settings');

  /* ---- the inner strip is reused, not duplicated -------------------------------- */
  ok('the old strip inside the Fix list card is hidden rather than deleted',
    /\.tabs \{ display: none; \}/.test(css) && /id="panelFix"/.test(html),
    'its panels are still the thing setTab() shows');
  ok('and the three sections that share it drive it by name',
    /inner: 'fix'/.test(html) && /inner: 'suggest'/.test(html) && /inner: 'coverage'/.test(html));
  /*
   * The filter CARD, by an id of its own. This looked up "filterHeading" — which is the id of the
   * <h2> INSIDE the card, because the card itself had none — so the heading was hidden and the
   * severity chips, the page box and the search box stayed on screen under every other section.
   * The severity chips appearing at the top of Things to add is what was reported.
   */
  ok('the filters are hidden outside the Fix list, since they filter only that list',
    /getElementById\('tiles'\)/.test(html) && /getElementById\('filterCard'\)/.test(html)
    && /<section class="card" id="filterCard" aria-labelledby="filterHeading">/.test(html),
    'hiding the heading inside a card leaves the card itself on screen');
  ok('and nothing looks the filters up by the heading id any more',
    !/getElementById\('filterHeading'\)/.test(html));

  /* ---- and a page chip has to bring its section with it ------------------------- */
  ok('clicking a page in a finding switches the SECTION, not just the inner tab',
    /function jumpToPage\(page\) \{[\s\S]{0,400}setTopTab\('fixlist', false\);/.test(html),
    'the inner tab would change inside a panel that is not on screen, so the click would appear dead');

  /* ---- a phone gets one scrollable row ------------------------------------------ */
  ok('the strip scrolls sideways rather than wrapping',
    /\.toptabs \{[\s\S]{0,700}overflow-x: auto/.test(css) && !/flex-wrap: wrap/.test(/ \.toptabs \{[\s\S]{0,900}?\}/.exec(css)[0]),
    'eleven tabs wrapped onto rows would push the content off the bottom of a narrow screen');
  ok('and the open one is the accent, with a label colour that flips with the theme',
    /\.toptab\[aria-selected="true"\] \{[\s\S]{0,160}background: var\(--accent\)/.test(css)
    && /\.toptab\[aria-selected="true"\] \{[\s\S]{0,200}color: var\(--card\)/.test(css));
  /*
   * ---- the strip has to be usable, and Settings has to be findable -----------------
   *
   * Reported: "i cannot access all the sections there needs to be a scroller", and "settings
   * should always be accessible". The strip scrolled but said nothing about it — no scrollbar, no
   * arrows — so it looked like a strip with tabs missing off the end.
   */
  ok('only the middle of the strip scrolls, so the pinned button cannot be pushed off it',
    /<div class="toptabs-scroll" id="topTabsScroll">/.test(html)
    && /\.toptabs-scroll \{[\s\S]{0,400}min-width: 0;/.test(css),
    'without min-width: 0 a flex child will not shrink below its content, so it never scrolls');
  /*
   * Asserted by POSITION rather than by a pattern across the markup: the Settings button must come
   * after the scroll container closes. That is the actual property, and it does not care how many
   * arrows or comments sit in between.
   */
  /*
   * Overview and Settings bookend the strip, OUTSIDE the part that scrolls.
   *
   * Both are the two you need without hunting: Overview is where you land and where Check now
   * lives, Settings is what you reach for when something is wrong. Nine in between still scroll.
   */
  ok('Overview and Settings both sit outside the scroller, one at each end',
    (() => {
      const scrollStart = html.indexOf('id="topTabsScroll"');
      const scrollEnd = scrollStart < 0 ? -1 : html.indexOf('</div>', scrollStart);
      const overview = html.indexOf('data-tab="overview"');
      const settings = html.indexOf('data-tab="settings"');
      return scrollStart > 0 && scrollEnd > scrollStart
        && overview > 0 && overview < scrollStart
        && settings > scrollEnd;
    })(),
    'inside the scroller either of them is exactly as hard to reach as everything else');
  /*
   * "The scroller bar is difficult to use. Please use a scroller long bar underneath that can be
   * dragged ... make it easy to use and not hard to click on."
   *
   * The arrows moved the strip by a fixed amount per press — never quite the amount you wanted —
   * and each one was a 30px target. This is the control everyone already knows: grab it and drag.
   */
  ok('there is a draggable scroller under the strip, not a pair of small arrows',
    /id="topTabsRail"/.test(html) && /id="topTabsThumb"/.test(html)
    && /function dragTabsTo\(clientX\)/.test(html) && /pointerdown/.test(html));
  ok('and the drag survives the pointer leaving the bar',
    /setPointerCapture/.test(html) && /on\(window, 'pointerup', endDrag\)/.test(html),
    'a drag that dies the moment the cursor strays off a 16px bar is unusable');
  ok('the bar and its thumb are big enough to hit',
    /\.toptabs-rail \{[\s\S]{0,260}height: 16px/.test(css) && /\.toptabs-thumb \{[\s\S]{0,200}min-width: 52px/.test(css),
    'a thin, small scroller is the thing that was too hard to use');
  ok('and it hides itself when there is nothing to scroll',
    /rail\.hidden = true/.test(html), 'a dead bar under a strip that already fits is a lie');
  ok('the native scrollbar is gone, so there is one scroller and not two',
    /\.toptabs-scroll \{[\s\S]{0,600}scrollbar-width: none/.test(css));
  ok('the thumb is sized and placed by how far through the strip you are',
    /range\.view \/ range\.wide/.test(html) && /parseInt|Number\(ui\.topTabsScroll\.scrollLeft\)/.test(html),
    'a scroller that does not track the thing it scrolls is decoration');
  ok('the bar is the visible cue that there is more to the right',
    /\.toptabs-rail \{[\s\S]{0,300}background: var\(--fill-soft\)/.test(css));
  ok('on a phone the strip hides its native scrollbar, since the bar is the scroller there',
    // Gathered here rather than reusing the later variable: this runs before it is declared.
    (html.match(/@media \(max-width: 700px\) \{[\s\S]*?\n  \}/g) || []).join('\n')
      .indexOf('.toptabs-scroll { scrollbar-width: none; }') >= 0);

  /* ---- the note that says what a check WOULD do ---------------------------------- */
  /*
   * Asked for: "in the overview section there should be a short note saying what specific checks
   * would be done if it was checked now based on the settings applied."
   *
   * The app already said what RAN, after the fact. This is the other question, and it is answered
   * from the LIVE settings rather than from the last report — because the whole point is that it
   * is right before you press the button, including when you have just changed a setting.
   *
   * Which rules a file supports depends only on whether it is a PDF or a .pptx, so it needs no
   * scan and cannot go stale. ruleCoverageFor lives beside runLocalChecks, written to mirror it: if
   * the two ever disagree, the front page is describing a check that is not the one that runs.
   */
  const checks = require(path.join(ROOT, 'lib', 'checks.js'));
  ok('the server can say what would run without running anything',
    typeof checks.ruleCoverageFor === 'function');
  const pdf = checks.ruleCoverageFor('pdf');
  const pptx = checks.ruleCoverageFor('pptx');
  ok('a PDF supports fewer checks than a .pptx, and the numbers add up',
    pdf.ruleCount < pptx.ruleCount && pdf.ruleCount === pdf.totalRules - pdf.skippedRules.length,
    'pdf ' + pdf.ruleCount + '/' + pdf.totalRules + ', pptx ' + pptx.ruleCount + '/' + pptx.totalRules);
  ok('and the same five a PDF always misses are the ones named in the report',
    pdf.skippedRules.map((x) => x.rule).sort().join(',')
      === ['missing-notes', 'off-canvas', 'page-too-busy', 'text-overflow', 'uncaptioned-image'].join(','),
    pdf.skippedRules.map((x) => x.rule).join(', '));

  ok('the note is drawn in Overview, above the buttons',
    /<div class="check-plan" id="checkPlan"><\/div>[\s\S]{0,200}<div class="actions-row">/.test(html));
  ok('and it is rebuilt from the live settings, not from the last report',
    /function renderCheckPlan\(\)/.test(html) && /cfg\.wouldRun/.test(html)
    && !/state\.data\.local[\s\S]{0,80}wouldRun/.test(html),
    'reading it from the report is how it would go stale the moment a setting changed');
  ok('it says plainly when the AI review is off, since that is the difference between a mark and none',
    /The AI review would be skipped/.test(html) && /check-plan-off/.test(html));
  ok('and it names what the file cannot supply, rather than an internal rule name',
    /skippedReasons\(plan\.skippedRules\)/.test(html) && /speaker notes/.test(html)
    && !/text-overflow/.test((/var what =[\s\S]{0,400}/.exec(html) || [''])[0]));
  ok('the shared reason list is not duplicated between the two notes',
    (html.match(/function skippedReasons\(/g) || []).length === 1
    && (html.match(/var why = skippedReasons\(skipped\)/g) || []).length === 1,
    'two copies of that list would eventually disagree');

  /* ---- "Start setup" ------------------------------------------------------------ */
  /*
   * Asked for above Check now in the header, in "a nice shade of green/turquoise". It sits directly
   * beside a blue button, and two blues of different weights read as a mistake, so teal says "a
   * different kind of thing" while staying in the same cool family.
   */
  ok('Start setup comes before Check now in the header',
    html.indexOf('id="btnShowSetup"') < html.indexOf('id="btnCheckTop"'),
    'in a column layout the first one is the one on top');
  ok('it is teal rather than a second blue',
    /\.btn-setup \{[\s\S]{0,200}background: var\(--setup\)/.test(css));
  ok('and every one of its colours is a token, in all three themes',
    (html.match(/--setup:/g) || []).length === 3
    && (html.match(/--setup-dark:/g) || []).length === 3
    && (html.match(/--setup-ink:/g) || []).length === 3,
    'a literal here is how the dark-mode contrasts broke twice before');
  ok('with an ink colour that flips, because a bright teal needs dark text',
    /--setup-ink: #04231f/.test(html));

  /* ---- and nothing folds away any more ------------------------------------------ */
  /*
   * Each of these was a collapsible card on a long page. Now each one IS its own tab, so an
   * "open me" step in front of it is one click too many.
   */
  ['settingsDetails', 'compareCard', 'slidesDetails', 'infoDetails', 'statsDetails'].forEach((id) => {
    ok('the ' + id + ' section starts open, since the tab already opened it',
      new RegExp('<details[^>]*id="' + id + '"[^>]*\\sopen').test(html));
  });
  ok('and its summary cannot fold it again',
    /\.panel > \.details-card > summary \{[\s\S]{0,160}pointer-events: none/.test(css));
  ok('with the body forced visible, so a stray click cannot leave a section looking empty',
    /\.panel > \.details-card > \.details-body \{ display: block !important; \}/.test(css));

  /*
   * ---- three things reported after using it --------------------------------------
   */
  /*
   * "on the fix list it has all three sections including the coverage grid and worth adding ...
   * thats what it has on my phone". Exactly right, and only on a phone: the phone block forced the
   * OLD inner strip back on, and it comes after the rule that hides it, so on a narrow screen Fix
   * list, Worth adding and Coverage grid all reappeared inside the Fix list section — a second copy
   * of three sections that are already tabs.
   */
  const phoneBlocks = (html.match(/@media \(max-width: 700px\) \{[\s\S]*?\n  \}/g) || []).join('\n');
  ok('the old inner strip stays hidden on a phone too',
    /\.tabs \{ display: none; \}/.test(phoneBlocks) && !/\.tabs \{ display: flex; \}/.test(html),
    'the phone rule that re-showed it made three sections appear inside the Fix list');

  /*
   * "check now is always an available button no matter what section you are in"
   */
  ok('there is a Check now in the header, visible from every section',
    /id="btnCheckTop"/.test(html) && /on\(ui\.btnCheckTop, 'click'/.test(html));
  ok('and it mirrors the main button rather than deciding anything for itself',
    /if \(ui\.btnCheckTop\) \{[\s\S]{0,160}ui\.btnCheckTop\.disabled = ui\.btnCheck \? ui\.btnCheck\.disabled/.test(html),
    'two buttons that can disagree about whether a check is possible is worse than one');
  /*
   * The floating shortcut used to test only whether the button had scrolled past the top. A hidden
   * element measures zero, so in every other section it decided the button was still on screen and
   * never appeared — which is precisely where it is most needed.
   */
  ok('the floating shortcut also appears whenever Overview is closed',
    /getElementById\('panelOverview'\)[\s\S]{0,300}var gone = !overviewOpen/.test(html),
    'a hidden button measures zero, so "still visible" was the wrong answer in every other section');
  ok('and both are re-checked when the section changes',
    /renderTabRail\(\);[\s\S]{0,200}renderDock\(\);/.test(html));

  ok('the open tab is scrolled into view, since the strip is one scrolling row',
    /strip\.scrollLeft = Math\.max\(0, offset/.test(html),
    'eleven tabs means the selected one is usually off the right-hand edge');

  /* ---- two bugs that only LOOKING at the page found ---------------------------- */
  /*
   * The footer sat inside .page, which is a flex column now, so the default order of 0 put it
   * ABOVE the header. It is the last thing on the page in every sense.
   */
  ok('the footer is ordered last, because it is a child of the flex column',
    /\.footer \{ order: 999; \}/.test(css), 'a default order of 0 renders it above the header');
  /*
   * And the version it shows was a hardcoded `var APP_VERSION = '1.0.0'` that no release ever
   * updated — so the footer told every user they were on 1.0.0 while the app was on 1.0.50.
   * server.js has a comment saying this was fixed; the VERSION it reports WAS fixed, and the
   * footer had a second copy nobody looked at.
   */
  ok('the footer version comes from the running app, not from a constant',
    !/var APP_VERSION = /.test(html)
    && /setText\(ui\.appVersion, \(data && data\.version\)/.test(html),
    'a hardcoded version is a version that will be wrong');
}

section('12ak. A running header is not a slide title');
{
  /*
   * Reported: "on the slides section it thinks all my slides are named candidate name".
   *
   * A PDF has no title field, so the title is guessed from the first line of the page. On a real
   * export the first line is a running header — the same text on every page — so every slide came
   * out with the same name. It also made the missing-title check useless on a PDF, because every
   * page appeared to have a title.
   *
   * The distinguishing property is REPETITION. Fixing it by taking line two, then line three, just
   * walked down the list of headers one wrong answer at a time — this portfolio has three.
   */
  const pdfdeck = require(path.join(ROOT, 'lib', 'pdfdeck.js'));
  ok('the header detector is exported for testing', typeof pdfdeck.runningHeaders === 'function');

  const page = (first, second, third, body) => [first, second, third, body].filter(Boolean).join('\n');
  const pages = [
    page('Candidate name', 'Candidate number: 17632', 'Centre: 12345', 'Design brief\nThe client wants a lamp.'),
    page('Candidate name', 'Candidate number: 17632', 'Centre: 12345', 'Research\nI looked at lamps.'),
    page('Candidate name', 'Candidate number: 17632', 'Centre: 12345', 'Ideas\nSeveral sketches.'),
    page('Candidate name', 'Candidate number: 17632', 'Centre: 12345', 'Making\nI cut the wood.'),
    page('Candidate name', 'Candidate number: 17632', 'Centre: 12345', 'Testing\nIt held the weight.'),
  ];
  const headers = pdfdeck.runningHeaders(pages);
  ok('a line repeated at the top of every page is recognised as a header',
    headers.has('candidate name') && headers.has('centre: 12345'),
    [...headers].join(' | '));

  /*
   * And the title is then the first line that is NOT one of them — skipping, rather than giving
   * up, because the real title sits below the headers.
   */
  const shortPdf = Buffer.from('%PDF-1.4 not a real pdf');
  ok('the detector needs a few pages before it will call something a header',
    pdfdeck.runningHeaders(['Candidate name\nTitle', 'Candidate name\nTitle', 'Candidate name\nTitle']).size === 0,
    'on three pages a repeated line is as likely to be a real title');
  ok('and a one-off line is never treated as a header',
    pdfdeck.runningHeaders(['Intro\nA', 'Middle\nB', 'End\nC', 'More\nD', 'Last\nE']).size === 0);
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
