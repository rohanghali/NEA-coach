# CONTINUE HERE — read this first

**For any AI assistant picking this project up, and for Rohan.**

This file is the complete handover. Read it before changing anything. It records
what the app does, the rules that must not be broken, and what is still to do.

Last updated by: the session that built the app (DeepSeek Harness on a MacBook).
The app was built across one long session and is working and tested.

---

## 1. What this is

**NEA Coach** — a read-only reviewer for a UK A-level Design and Technology NEA
portfolio. It reads a `.pptx` file, finds problems worth fixing, ranks them by
priority, and reports how well the pages evidence each mark-scheme criterion.

It was built for Rohan, who is doing the **WJEC Eduqas A-level Design and
Technology** non-exam assessment (NEA), a portfolio of roughly 200 A3 PowerPoint
pages. His GCSE portfolio (60 pages, AQA, scored 99/100) is used as the test case.

Runs entirely on his own Mac. No server, no subscription, no accounts.

---

## 2. Where it lives and how to run it

```
NEA-Coach/
├── Start NEA Coach.command   <- double-click this on a Mac (right-click > Open the FIRST time)
├── Start NEA Coach.bat       <- double-click this on Windows (see section 7: never run on a real PC yet)
├── server.js                 <- local web server, scan orchestration, report building
├── lib/
│   ├── zipread.js            <- minimal read-only ZIP reader
│   ├── pptx.js               <- OOXML parser (text, fonts, images, autofit, layout inheritance)
│   ├── imagemeta.js          <- image dimensions from file headers
│   ├── checks.js             <- the free, offline, deterministic rules
│   └── ai.js                 <- Gemini client, batching, cache, cost accounting
├── public/index.html         <- the ENTIRE frontend, self-contained (no build step)
├── spec/                     <- mark scheme files. Loaded as the criteria to judge against
├── sample/                   <- a 24-page test deck with deliberately planted flaws
├── state/                    <- config.json, cache.json, cost.json, last-report.json
└── runtime/                  <- Node binary, downloaded on first run if needed
```

The portfolio `.pptx` sits **one level up**, next to the `NEA-Coach` folder.
`config.json` uses `"deckDir": ".."` (a relative path) on purpose, so the app keeps
working after being copied to a different Mac with a different user name.

Then open <http://127.0.0.1:4173>.

---

## 3. RULES THAT MUST NOT BE BROKEN

These were deliberate decisions. Changing them will break the app or break trust.

1. **NEVER write to the `.pptx`.** The whole point is that portfolio regulations
   require the work to be the student's own. `lib/zipread.js` opens the file with
   `fs.readFileSync` only. There is no write path. This was verified by hashing a
   527 MB file before and after a full scan: identical SHA256 and mtime.
   If you add a feature, do not break this.

2. **ZERO third-party dependencies.** No `package.json`, no `node_modules`, no
   build step. Only Node built-ins (`fs`, `path`, `http`, `https`, `crypto`,
   `zlib`). This is why it runs unchanged on Intel and Apple Silicon, and why it
   cannot be broken by a supply-chain change. If you think you need a package,
   you almost certainly do not.

3. **Do not show a mark-like score from the free checks.** See section 6. The free
   checks measure presentation and completeness. They cannot grade design
   thinking. An earlier version showed "51 out of 100" for a portfolio that
   actually scored 99/100, which was indefensible.

4. **The portfolio is a WORK IN PROGRESS.** Never suggest adding something that is
   already present. This is why AI findings carry `kind` and `progressAware`.
   Telling a student to "add an evaluation" when they have written three pages of
   one destroys their trust in the whole report.

5. **Only touch this folder.** Everything the app writes goes in `state/`.

6. **Tables and images are CONTENT, not decoration.** A `.pptx` table
   (`<a:tbl>`) holds real text that must be extracted like any other text. It was
   once detected but not read, so four complete manufacture-plan pages reached
   the AI as two words each and it told the student to write pages he had already
   written. If you add a shape type, make sure its text is extracted.

7. **Never call a page blank.** A page with images and little text is a
   picture-led page, which is normal and often correct. The prompt states
   explicitly that such pages must not be reported as blank. Only a page with no
   text AND no images is genuinely empty.

8. **The fix list is ordered by relevance, not by severity.** Local checks carry a
   `tier` (`substance` / `presentation` / `obvious`) and AI findings are
   `substance`. The UI shows substance first and folds mechanical checks away.
   Do not reorder this back to severity-first: see rule 9.

9. **Relevance must outrank everything in the sort.** `sortForFixList` once sorted
   on `rank` before `priority`. Local checks are numbered from 1 and AI findings
   have no rank (null), so all ~100 mechanical checks were forced above every
   piece of AI judgement. The student's words: *"the list of 'fix it' stuff seems
   to still only be the silly stuff like too much text."* That was this bug.

10. **Cost must be visible before it is spent.** Prices are per model and shown in
    the picker and above the Check button. Never add a paid feature without
    saying what it costs. The student's key pays for it, not ours.

11. **Only this Mac gets in without the access code.** The server listens on the
    network for the phone, so every other device needs `?key=` (the
    `accessCode` in `config.json`), the Host header is checked, and every POST
    must carry `X-NEA-Coach: 1`. Never return `accessCode` or `apiKey` from an
    endpoint; use `publicConfig()`. Tests: section 10.

12. **The AI reads 12 pages at a time, so batches cannot judge the whole
    portfolio.** Batch "add" findings are only candidates. One cross-check
    request per scan (`buildCrossCheckPrompt` in `lib/ai.js`) sees an outline of
    every page and decides each candidate: missing, improve-existing (moved to
    the fix list), already-done or duplicate. It also sets whole-portfolio
    coverage. If it fails, additions are held back and coverage falls back to
    the STRONGEST section report, never the weakest. Tests: sections 11 and 11b.

13. **One exam board at a time.** `spec/` holds a criteria file per board
    (`eduqas-…` and `aqa-…`). `loadSpec(specDir, specId)` loads only the selected
    board and reports the other board's file as skipped. Loading both would send
    two mark schemes on every paid request and mix the boards in one coverage
    grid. Changing board in Settings clears the cached AI judgement
    (`cache.batches` and `cache.crossCheck`) because those were made against the
    other board's criteria. Tests: section 11d.

14. **The predicted mark is summed by the app, never by the model.** The model is
    asked for a `suggestedMark` per criterion; `summarisePredictedMarks` clamps
    each to that criterion's published weight and adds them, so the total cannot
    exceed 100%. A criterion the model skipped falls back to its coverage
    percentage and is flagged `estimated`. Both boards now have real published
    weights (Eduqas 15/15/25/25/20, AQA 10/10/20/20/20/20), so the equal-shares
    path is a safety net rather than the normal case: it still exists for a board
    whose file has no weights, and `weightsPublished: false` makes the UI say so.
    Tests: sections 11c and 11d.

15. **A prompt change must bump `PROMPT_VERSION` in `lib/ai.js`.** The batch cache
    key is `model + specHash + fingerprints`, and specHash includes
    `PROMPT_VERSION`. Without a bump, editing the prompts leaves every
    already-read page serving its OLD findings from the cache, so a prompt
    improvement looks like it did nothing. It also silently keeps the previous
    prompt's coverage and marks. Bump it whenever the prompts, the response
    schema, or the derivation of findings changes.

16. **The visual review is the accurate judge; the text pass produces the detail.**
    `runVisionPass` sends the portfolio PDF as ONE `inlineData` part and asks for
    whole-portfolio coverage plus a mark per criterion. It runs ONCE per scan, not
    once per batch: sending the PDF with all five batches would re-send every page
    five times. Its coverage beats the text-only cross-check's (guarded inside
    runAiPass), because a text-only judge cannot see a sketch, and its findings are
    folded in with `+1500` priority because they describe the actual page images.
    Measured on a real 60-page A3 deck: about 34,000 input tokens for a whole-deck
    call, under 1.2p. With no PDF stored the pass is skipped entirely. Tests: 11e.

16b. **MEASURED: the visual review is stable to about ±1 mark, so `VISION_RUNS`
    is set for safety, not necessity.** Four repeats of the identical portfolio
    with the identical config gave totals of 82, 81, 82, 80 — mean 81.3, spread 2,
    standard deviation 0.8. Criteria a, b and c were identical in all four runs;
    only d, e and f moved, by 1 to 3 marks each.

    This retires an earlier claim in this file and in the wishlist that "the app's
    own noise is ±7 points". That was WRONG: the 88% and 81% figures came from two
    DIFFERENT prompt versions (before and after the suggestions field was added),
    not from two runs of one version. Comparing them measured the prompt change,
    not the noise. Do not repeat that mistake — only repeat runs of the SAME
    configuration measure variance.

    The implication to keep in mind: the run-to-run precision is already good, so
    the outstanding problem is ACCURACY — 81.3% consistently on a portfolio whose
    true mark is higher. A consistent error cannot be fixed by averaging; it needs
    an anchor (see section 7).


17. **A text-only prediction under-scores design work, and the app must say so.**
    The same real portfolio scored 99/100 and was predicted at 75% text-only and
    88% once the pages were visible. Its most visual criterion — brief and
    specification — moved 6/10 to 9/10. Whenever a predicted mark is shown, the UI
    states whether the AI actually looked at the pages, so a text-only figure is
    never mistaken for a comparable one. Do not remove that wording.

18. **Never quote a text-only price while sending images.** Both
    `estimateScanCostGbp` (server) and `estimateScanGbp` (page) take the stored PDF
    into account, and their token constants must stay in step — a test asserts the
    vision estimate is higher and still under 10p for 60 pages. The same applies to
    the MARK TOGGLE: `visionRunsFor()` in `lib/ai.js` is the single source of truth
    for how many runs happen, mirrored by `VISION_RUNS_MIRROR` in the page. These
    drifted once — the page halved its quote when marks were switched off and the
    server did not, over-quoting by three times.

19. **A finding's `id` is NOT stable across scans; never diff on it.**
    `makeFinding` appends a global sequence number (`wall-of-text@7#29`, then
    `#55`), and an AI finding's id embeds its title as well. Comparing raw ids in
    `compareReports` would report EVERY finding as both fixed and new — a diff
    that looks like it works and is entirely wrong. Use `findingKey()`: rule plus
    page for a local check, normalised title for an AI finding. Tests: 11g.

20. **The dark theme redefines variables only, and the palette is declared once.**
    The `--dark-*` tokens are the single definition; both the system rule
    (`@media (prefers-color-scheme: dark)`) and the explicit rule
    (`:root[data-theme="dark"]`) map onto them. Writing the palette twice was how
    the two would drift apart. `data-theme` is REMOVED for "system" so the media
    query governs, and set for light/dark so an explicit choice wins. A test-ish
    check worth repeating after any CSS change: every `var(--x)` must have a
    definition, and the two dark blocks must set identical keys.

21. **CSS variables that are referenced but never defined fail silently.**
    `color: var(--text)` had no fallback anywhere, so it produced no colour at all;
    the `--warn-*` trio only had light fallbacks, leaving pale amber panels on a
    dark background. Both are fixed. If you add a variable, define it in the light
    root AND both dark blocks, or give it a fallback.

---


## 3b. Reading an exam-board specification

Two traps worth knowing before you try to add a third board.

**AQA's specification PDFs are permission-locked.** `AQA-8552-SP-2017.PDF` is
encrypted (`/Encrypt` with an empty *user* password), so it opens in any viewer
but its streams are RC4/AES encrypted. Hand-rolled `zlib` extraction cannot work
on it, and neither can the app, which has no PDF parser and must not gain a
dependency. The reliable route on a Mac is CoreGraphics, which decrypts it
transparently:

```bash
osascript -l JavaScript -e '
ObjC.import("Quartz");
var d = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath("/path/to/spec.pdf"));
console.log(ObjC.unwrap(d.string));' > /tmp/spec.txt
```

**Do not confuse the exam with the NEA.** The AQA *written exam* is also 100
marks, split 20/30/50 across sections A, B and C. The NEA criteria are a
different thing entirely (10/10/20/20/20/20 in section 4.4.4). Taking the exam
split as the criteria weights would silently give every AQA student a wrong
predicted mark.

---

## 4. How to test a change

There IS a test suite now. Run it first, always:

```bash
cd NEA-Coach
node tests/test.js      # 547 assertions, no dependencies
```

It covers the bugs that actually happened: table extraction, disabled image
rules, the sample answer key, relevance tiers, AI classification, pricing and
longest-prefix matching, prompt construction, key secrecy, read-only safety, the
predicted-mark maths, and keeping the two exam boards separate.
**Add an assertion for every bug you fix.** Each existing test exists because
something broke and a student noticed.

**Golden rule: always test against BOTH decks.** The sample deck is tidy and
catches logic errors; the real 60-page deck catches everything else. Several bugs
were invisible on the sample. A useful smoke test on real data:

```bash
cd NEA-Coach
node -e "
const {parsePresentation}=require('./lib/pptx');
const {runLocalChecks}=require('./lib/checks');
const d=parsePresentation('../Rohan Ghali.pptx');
const r=runLocalChecks(d);
console.log('findings', r.findings.length, 'byTier', JSON.stringify(r.byTier));
console.log('errors', JSON.stringify(r.errors));
"
```

Then the manual syntax pass:

```bash
cd NEA-Coach

# 1. Syntax check everything
for f in server.js lib/*.js tests/*.js; do node --check "$f"; done

# 2. Run the free checks against the sample deck, which has KNOWN planted flaws
node -e "
const {parsePresentation}=require('./lib/pptx');
const {runLocalChecks}=require('./lib/checks');
const d=parsePresentation('sample/AQA-GCSE-Portfolio-SAMPLE.pptx');
const r=runLocalChecks(d);
console.log(r.findings.length, 'findings, score', r.score, 'errors', JSON.stringify(r.errors));
r.findings.forEach(f=>console.log(' ', f.rule, 'p'+f.slide));
"
```

`sample/PLANTED-FLAWS.txt` is the answer key. The checks must keep catching:
no title on pages 4/13/21, low-res images on 6/7/14, duplicate pages 11+12,
near-empty 16/17, uncaptioned 6/7 (page 14 has a caption and must NOT be flagged).

**Always test against BOTH decks** — the tidy sample and the messy real one. Several
bugs only appeared on the real portfolio and were invisible on the sample:

- Font sizes are usually **inherited from the slide layout**, not stored on the
  slide. Ignoring that made the "text too small" check silent on real work.
- Run properties are written **self-closing** (`<a:rPr .../>`), which a
  `<tag>...</tag>` matcher cannot see. That lost font sizes AND text.
- `<a:extLst>` inside `<p:cNvPr>` looks like `<a:ext>` to a loose regex, so
  geometry must be read from inside `<a:xfrm>` only.
- The real portfolio captions images **beside** them, not below.
- `<a:tbl>` table text lives in `<a:tc><a:txBody>`, and the rows/columns must be
  walked as DIRECT children (`childElements`), not with `elements()`, which
  flattens the grid.
- A table's `<p:xfrm>` height is NOT a clipping boundary: PowerPoint grows rows to
  fit. Estimating overflow from it produces a false positive on every table page.

### Gemini API traps (all verified against the live API, all cost money if wrong)

- **Newer models reject `thinkingConfig` outright** with HTTP 400 "Request
  contains an invalid argument": `gemini-3.5-flash-lite` and
  `gemini-flash-lite-latest`. Sending the field at all fails the request, even as
  `thinkingBudget: 0`. See `THINKING_UNSUPPORTED` in `lib/ai.js`; `callWithRetry`
  also drops the field and retries if a model ever refuses it.
- **`listModels` returns retired models** that then 404 with "no longer available
  to new users" (`gemini-2.5-*`, `gemini-2.0-*`). See `RETIRED_MODEL`. They look
  like the cheapest options, so offering them is a trap.
- **Output price includes thinking tokens** on every Gemini 3 model. A model left
  "thinking" can cost several times the naive estimate.
- **Longest-prefix price matching is essential.** A plain `find()` matched
  `gemini-3.8-flash` against `gemini-3.8-flash-lite` and overcharged the meter
  about sevenfold.
- **Published prices change** (see https://ai.google.dev/gemini-api/docs/pricing).
  The table in `lib/ai.js` was read from that page and must be re-checked rather
  than guessed. Batch and Flex tiers are 50% cheaper if cost ever needs cutting
  further.
- **503 "model overloaded" is common** and needs exponential backoff WITH jitter.
  Several batches retrying in lockstep kept failing together.

### Model selection traps

- **Never trust a stored model name.** Google retires models but keeps listing
  them, and a config pinned to one (the student's was on `gemini-2.5-pro`) makes
  every request 404 with nothing on screen to explain it. `pickModel` ignores a
  retired preference, `POST /api/config` refuses to save one, and `performScan`
  self-heals a config it finds naming one.
- **Do not cap or hide the model list.** It was a `<datalist>` limited to twelve
  entries, which is why the model the student wanted "isnt showing up in the
  options now". It is a `<select>` listing every usable model, no cap. A datalist
  also shows nothing at all until you type, so it cannot be browsed.
- **Filter out non-text models.** Image, TTS, music, robotics, Omni, transcribe
  and deep-research models all appear in `listModels` and none can do this job.
- **Prefer a thinking-capable model.** `gemini-flash-lite-latest` is cheap but
  rejects `thinkingConfig`, so picking it silently removes the deep-thinking
  option; `gemini-3.1-flash-lite` costs the same and supports it.
- **Keep the two price estimates in step.** The page mirrors the server's batch
  size and token constants, and a test asserts they still match. When they
  drifted, the app quoted 3.13p in one place and 3.35p in another.

```bash
# 3. Exercise the AI layer without spending money (pure functions, no network)
node -e "
const ai=require('./lib/ai');
// parseBatchResponse, normaliseAiFinding, pickModel, costFor are all exported
" 
```

---

## 5. What works today

**Free local checks (17 rules, offline, ~1 second for 200 pages):** missing titles,
overloaded pages, wall-of-text paragraphs, text too small, genuinely blocky images
(under 72 DPI), images shrunk too small to read, uncaptioned images, near-empty
pages, duplicate pages, text overflowing its box, elements off the printable page,
text auto-shrunk to fit, inconsistent heading sizes, UK/US spelling clashes, common
misspellings.

**AI review (needs a Gemini API key):** judgement against the mark scheme, coverage
grid, wording/tone/placement critique, and a separate "Worth adding" list.

**UI:** three tabs — Fix list, Worth adding, Coverage grid. Sortable slides table,
severity filters, search, localStorage-backed "done" checkboxes, live cost meter.

**Cost:** measured at roughly **£2.90 per year** of heavy use (150 incremental
scans plus 10 full rescans of a 200-page deck). Fingerprint caching means unchanged
pages are never re-sent.

---

## 6. The score question, explained once

`lib/checks.js` returns `presentationTidiness` (0-100), with a deprecated alias
`score`. **It is not a mark and must never be displayed as one.** It only colours
the progress ring. The UI shows a plain count of *points to check* instead.

A real AI mark estimate should be shown separately, labelled as an estimate against
the Eduqas criteria, and only once the AI has actually run.

---

## 7. Still to do

### Already built, waiting on a Gemini API key
Nothing further. The AI pass is written and unit-tested against realistic and
malformed responses. It has **never made a live call**, because no key was ever
available. `SETUP-AI.md` explains how to add one (free tier, no card).

### Confirmed unfinished

- ~~**iPhone / mobile view.**~~ **DONE.** The same web interface now works on a
  phone. The server binds `0.0.0.0`, the launcher prints the phone address
  (`http://<lan-ip>:4173`, or a `100.x` Tailscale address for access anywhere),
  and under 700px the Fix list shows the top 5 with "Show more", the filters fold
  away, and the layout tightens. Scanning works from the phone because it happens
  on the Mac. The API key never leaves the Mac. Verified at 428x926 (iPhone 12
  Pro Max): no horizontal scroll, 7/7 layout checks pass.
- **Follow-up items on skipped pages.** The app records `ai.specSkipped`, but
  there is no UI surfacing it.
- **Visual checking.** `checks.js` infers page layout from XML. Rendering slides to
  images and letting the AI look at them would catch things geometry cannot (poor
  contrast, cramped spacing, overlapping elements). This was prototyped: **the
  bundled LibreOffice Kit fails on the 527 MB deck** ("input too large"), and its
  direct PPTX-to-PNG path returns blank pages, so it must go PPTX to PDF to PNG.
  It also refuses to write to an output path that already exists.

### The phone address moves, and that looked like a bug

Reported as "the mobile app is not working - Safari cant open page". It was not a
layout bug and not a code bug: the server was running, bound to `0.0.0.0`, the
firewall was off, and `http://<lan-ip>:4173/?key=...` returned HTTP 200. **The
Mac's LAN address had changed** (the README example said `192.168.0.87`; the Mac
was on `192.168.1.136`), so the phone's old bookmark no longer resolved. An
address change is indistinguishable from a broken app unless you know to suspect
it.

Fixed properly rather than by telling the student to read the terminal again:
`GET /api/connect` returns the current addresses with the access code appended,
and **Settings -> Use it on your phone** displays them. That endpoint is the only
one that returns the code, so it refuses anything that is not this Mac - and it
refuses a LAN device that already holds a VALID code, because a device with the
code has no need to be given it. Tests: section 10.

Still worth improving: nothing detects the address changing while the app is
running, so the page could in principle notice and say so.

### Calibration — measuring accuracy (not yet done)

`tools/calibrate.js` exists to answer the one question the app cannot answer about
itself: **is the predicted mark close to the truth?** It compares the app's
per-criterion marks against marks a teacher or moderator actually awarded.

```
node tools/calibrate.js run <pdf> --marks a=10,b=10,c=20,d=20,e=20,f=19
node tools/calibrate.js list
node tools/calibrate.js show <id>
```

It reports the per-criterion difference, the signed mean bias, the spread around
it, the worst criterion, and a verdict on the SHAPE of the error, because the
shapes need different fixes: a uniform offset is recalibratable, an uneven error
is not. Results append to `state/calibration.json`. One run costs a few pence.

**Blocker: no anchor exists yet.** What has been established, so there is no need
to redo it:

- The 99/100 AQA GCSE portfolio (`Rohan Ghali.pdf`, 60 pages) is present and IS
  the portfolio that genuinely scored 99. It is the file every vision measurement
  has used.
- **Its text contains no marks.** The whole PDF was extracted with CoreGraphics
  (`osascript -l JavaScript` + `PDFDocument`) and searched: no mark, no total, no
  teacher annotation, no mark sheet. The cover page carries only candidate
  details. So the total (99) is known but the per-criterion split is NOT, and the
  split is what calibration needs.
- **What would unblock it:** the six awarded marks for that portfolio from a mark
  sheet, or a teacher marking ONE section of any portfolio against a known mark
  scheme. One anchored criterion would already reveal the direction and size of
  the bias.
- Do NOT tune prompts before that exists. Changing a prompt without an anchor
  changes the number and proves nothing about whether it moved towards the truth.

### MEASURED baseline: the app is consistently 18 marks low

Recorded in `state/calibration.json` (view it with `node tools/calibrate.js show cal-baseline-1`).
Against the AQA GCSE portfolio that really scored 99/100:

| | |
|---|---|
| App's mean prediction | **81** (runs: 82, 81, 81) |
| Real total | **99** |
| Shortfall | **18 marks** |

Per criterion, the mean of three runs: a=9/10, b=9/10, c=17/20, d=17/20, e=14/20,
f=15/20.

Two hard-won lessons are baked into that record:

1. **An earlier 88 was an OUTLIER, not a regression.** Four runs of the current
   prompt gave 82, 81, 82, 80; three runs of the pre-suggestions prompt gave 82,
   81, 81. Both mean 81.3. So the prompt change made NO measurable difference, and
   the claim that it had regressed the marks (made in this file and to the user)
   was wrong — it compared a single lucky run against a mean. **Never judge a
   prompt change from one run when the spread is known.**
2. **Precision is good; accuracy is the problem.** Spread is under 1 mark across
   repeats, so this is a CONSISTENT 18-mark shortfall. Averaging more runs cannot
   fix a consistent error, and neither can prompt wording: something structural is
   making the app value this portfolio lower than its teacher did. The
   per-criterion split is what would show whether that is one criterion or all six.

Do NOT apply a +18 fudge factor to the mark. One anchor cannot tell an even
shortfall from a concentrated one, and inflating every portfolio's mark by 18
would wreck the app's accuracy on any portfolio that was not this one.

### THREE anchors: the cheap model COMPRESSES its scores

A third portfolio arrived (Jacob, AQA GCSE, 44 pages, real mark **89**). Measuring
it settled the shape of the error, because the three anchors now span a real range
of 89-99 rather than being bunched:

| Model | App | Real | Error |
|---|---|---|---|
| flash-lite | 81 | 99 | −18 |
| flash-lite | 80 | 96 | −16 |
| flash-lite | **66** | **89** | **−23** |

Fit across all three (`tools/calibrate.js fit`):

```
real ≈ app × 0.595 + 49.7      max residual 1.2 marks
```

**The slope is 0.6, not 1.** That is the whole story: flash-lite does not simply
mark low, it **compresses every portfolio towards the middle** — it gave a 99, a 96
and an 89 portfolio 81, 80 and 66. So the rubric it applies has far less range than
the mark scheme does, and a flat offset could never fix it: +18 would turn the 89
portfolio into 84, while the real answer needed +23.

Why this matters for the decision already made: `gemini-3.5-flash` gave **97** and
**92** raw, i.e. **no meaningful compression**, which is why it needs no correction
and remains the default. The compression is a property of the cheap model, not of
the prompt or the rubric.

A NOTE ON WHAT IS STILL TRUE: the 3-anchor fit is good (max residual 1.2), but it is
a correction for a model the app no longer uses by default. It is recorded as
evidence for WHY the default changed, not as something to apply.

### SOLVED: the accuracy problem was the MARKING MODEL, not the prompt

This is the most important finding in the project. Holding the PDF and the prompt
constant and changing only the model used for the marking runs:

| Model | Rohan (real 99) | Mithil (real 96) | Spread | Cost/check |
|---|---|---|---|---|
| `gemini-3.1-flash-lite` | median **81** (runs 59-93) | 70-80 | **~20 marks** | ~9p |
| **`gemini-3.5-flash`** | median **97** (96, 97, 98) | **92** | **2 marks** | ~16p |
| `gemini-3.8-flash` | median **92** (90, 92, 93) | **86** (86, 86, 87) | 1-3 marks | ~13p |

So the cheap default was **16-26 marks low with a ~20-mark spread**, while
`gemini-3.5-flash` is **within 2-4 marks on BOTH portfolios with a 2-mark spread**.
It fixes accuracy and repeatability at the same time.

`visionModel` therefore defaults to `gemini-3.5-flash` in `defaultConfig()` and in
`state/config.json`. Tests in section 10 assert it is not blank and not a
Flash-Lite, because "optimising" it back to the cheaper model would silently
return the app to being 20 marks wrong.

**What this means for the earlier work, honestly:**

- Five attempts at tuning the marking PROMPT moved nothing (and two made it
  worse). Prompt wording was never the lever.
- The ~17-mark shortfall that looked like a calibration problem was model choice.
  No fudge factor is needed, and none should be added.
- Flash-Lite remains fine for the FINDING-by-finding text pass, which is what the
  fix list is built from. Only marking needs the stronger model, and only marking
  pays for it. That is why `visionModel` exists separately from `model`.

### TWO anchors now exist: the shortfall is near-constant, about 17-18 marks

Both are AQA GCSE portfolios with real total marks, recorded in
`state/calibration.json` (see `node tools/calibrate.js fit`):

| Portfolio | App's median | Real | Shortfall |
|---|---|---|---|
| 99/100 | 81 | 99 | **18** |
| 96/100 (Mithil) | 80 | 96 | **16** |

Two anchors at different marks is the minimum that can tell a constant offset from
a proportional error, and these two say **near-constant**: 18 marks low on the
stronger portfolio and 16 on the slightly weaker one.

**Both anchors record the MEDIAN of the runs, not the mean**, because the median is
what the app itself reports. Recording the mean understates the app by a mark and
silently corrupts the fit — an easy mistake to repeat.

**The two anchors are only 1 mark apart on the app's scale** (81 and 80), which has
a consequence worth understanding: two points always fit a line exactly, and these
two imply a slope of ×3 — three real marks per app mark — which is an artefact of
them being bunched, not a relationship. `fitCorrection` detects this (`appRange < 5`)
and falls back to the flat offset, saying why, rather than handing over that slope.
A third portfolio that the app rates much lower or higher is what would make a
slope meaningful.

**But it is still NOT applied, deliberately:**

- Two points fit a line exactly, so the fit proves nothing about portfolios
  outside the 79-81 range the app produced. `fitCorrection` says this itself.
- Both anchors are high-scoring AQA GCSE portfolios that the app rates 79-81.
  A portfolio the app rates 40 might really be 58 (constant) or 49 (proportional);
  a flat +17 would wreck one of those.
- A third anchor at a genuinely different app score is what would settle it.

### Reliability: occasional low outlier runs are REAL, not parse failures

Mithil's PDF gave runs of 81, 80, then **62 and 71** on re-runs. Investigated:
`finishReason: STOP`, valid JSON, ~1,600 output tokens, no error — the model
genuinely marked it lower, with criterion (a) falling from 9/10 to 6/10 and the
status moving from "strong" to "partial". It is model variation, not a bug.

Comparing spreads on the same code:

| Portfolio | Runs | Spread |
|---|---|---|
| Rohan (60 pp, 445 images) | 82, 81, 81, 82, 81, 80 | ~2 marks |
| Mithil (60 pp, **566 images**) | 81, 80, 62, 71 | **19 marks** |

So the "spread under 1 mark" measured earlier was specific to one portfolio, not a
property of the app. The 3-run MEDIAN is what absorbs this — it gave 80 for Mithil
against a mean of 74.3 — which is the strongest justification yet for `VISION_RUNS`
existing. Do not reduce it to 1 on the basis of Rohan's numbers alone.

### There is NO mobile app, and that keeps causing confusion

Asked twice now ("have you done the mobile app", "the mobile app is not working").
The answer: **there is no separate mobile app and never was.** The phone opens the
SAME `public/index.html`, served by the Mac, reflowed by the one
`@media (max-width: 700px)` block in the stylesheet. So:

- "Fix the mobile app" = edit that media query, or fix the connection.
- The address the phone uses changes when the router reassigns it, and that looks
  exactly like a broken app. See the phone-address section above and
  `GET /api/connect`.

**The real mobile bug found this round:** every one of the seven UI elements added
in recent sessions had **zero** rules in that media query — including `.connect-url`,
a `<code>` containing an IP address, which is one long unbreakable token and a
guaranteed horizontal overflow at 428px. That is how a desktop-only addition
breaks the phone without anyone noticing.

Fixed, and made testable: test section 11i asserts that a listed set of elements
each has a phone rule, that `.connect-url` may break mid-token, and that the
predicted-mark figure stacks instead of sitting in two columns. Verified by
removing a rule and watching the test fail with `missing: connect-url`.

**What that test is NOT:** a layout check. It proves the rules exist, not that the
page looks right on a phone. Nobody has viewed this on a device since the recent
additions, and it should not be described as verified until someone does.

### Tried and REJECTED: adding awarding-body guidance to the marking prompt

Measured, on the 99/100 portfolio, three runs each:

| Vision prompt | Runs | Mean |
|---|---|---|
| Before | 82, 81, 81 | **81.3** |
| With awarding-body guidance added | 72, 79, 76 | **75.7** |

The guidance added was real and sourced (marks are awarded holistically; best fit
puts fully-met work at the TOP of a band; the criteria are not a linear sequence;
the expected evidence list; "do not penalise missing annotation where the board
does not require it"). **Adding it LOWERED the mark by 5.6 and made accuracy
worse.** It was reverted, and `PROMPT_VERSION` went to 7 so the cache re-judges.

Why it probably backfired: a longer, more detailed instruction list makes the
model enumerate more shortcomings, and enumerating shortcomings pushes the mark
down. Whatever the mechanism, the measurement is the authority — do not re-add
this without measuring it again the same way. **Prompt engineering is not the
lever for this gap.**

### WORKED: board-aware annotation policy (this was a real accuracy gain)

25 of 103 findings on the real portfolio told the student to caption images, but
WJEC/Eduqas states in its own GCE D&T FAQ that "annotation throughout the work is
not necessary" and that what is required is a detailed summative comment on the
DT2 marksheet. The app was giving advice against the board's own rules, and it
buried the advice that mattered.

`SPEC_PRESETS` in `lib/ai.js` now carries `annotationRequired` per board (false for
Eduqas, true for AQA), `loadSpec` passes it through, and `runLocalChecks(deck,
options)` passes it to the rules. `uncaptioned-image` returns immediately when the
board does not require annotation. Measured on the real deck: **103 findings become
78.** Tests cover both boards and assert the rule actually does not fire, not
merely that a flag is set.

### Four later additions (all tested)

1. **Spelling section.** `TYPOS` in `lib/checks.js` — about 40 clear-cut
   misspellings. Deliberately NOT included: anything arguable. `dependant` is
   correct British English for a person and `publically` is an accepted variant, so
   both were removed; `millimetres`, `aluminium` and `medieval` are correct D&T
   vocabulary and must never be flagged. A confusing-pairs rule
   (affect/effect, loose/lose) was written, MEASURED, and deleted: it flagged 9
   pairs on the sample deck alone, which is noise, not help. Both words are real,
   so a rule that guesses is wrong too often.
2. **Per-page change tracking.** `updatePageHistory()` in `lib/checks.js`, kept in
   `cache.pageHistory`. PowerPoint has no reliable per-slide date — it rewrites the
   archive on save — so the app keeps its own. Pages are matched by CONTENT, so
   inserting a page marks one page new and the rest "moved", not "modified".
   Tests: section 11j.
3. **Scan history.** `state/history.json` holds up to 25 SUMMARIES (`snapshotOf`,
   1.9 KB against a 207 KB report). `GET /api/history` lists them;
   `?scan=N` returns one. The UI fetches on demand rather than shipping it with
   every page load. Tests: section 11k.
4. **Drag and drop.** In `public/index.html`: the whole page is a drop target,
   routed by file type (.pptx to the portfolio, PDF to the visual review, anything
   else refused by name with a banner). Verified only by code reading — nobody has
   dragged a file onto it in a browser.

### The meter shows ALLOWANCE, not money (free keys)

The student is on the free tier deliberately and will not add billing, so a money
meter is the wrong instrument — it reads £0.00 forever. Google's docs confirm the
mechanism (`ai.google.dev/gemini-api/docs/rate-limits`): limits are **requests per
minute, tokens per minute, and REQUESTS PER DAY**, measured **per project** rather
than per key, with the daily quota reset at **midnight Pacific**. For a free key the
cap that bites is RPD.

So the cost card is now **"Today's free AI allowance"**: percentage used, the
per-model split, and the time until reset. `state/quota.json` holds the counts.

Two things worth knowing:

1. **Google does not publish the per-model daily numbers.** The rate-limits page
   points at AI Studio for the live values, and the error body is what actually
   states the limit. So `ASSUMED_DAILY_REQUESTS = 20` is a starting point, and
   `learnQuotaLimit()` **replaces it with the real figure the first time a 429
   arrives** — the error carries `quotaValue` and the model. The UI labels an
   assumed limit "(estimated)" so a guess is never shown as a fact.
2. **The reset time is midnight PACIFIC, and getting it wrong is invisible.** A
   first attempt returned 15:00 UTC for "midnight Pacific" — 08:00 Pacific, eight
   hours early. It now SEARCHES for the instant that reads exactly 00:00 in that
   zone rather than deriving it from an offset, which is exact and cannot get a
   sign wrong. Tests: section 11o.

Request counting is threaded all the way to the HTTP call
(`onRequest` down through `runAiPass` → `callWithRetry` → `callGemini`, plus
`listModels` and the vision path), so retries and the model probe are counted. An
under-count would tell the student they had allowance they did not.

**Measured usage:** one full 60-page check is about **9 requests** (5 text batches
+ 1 cross-check + 3 marking runs), so the free tier allows roughly **two checks a
day**. Repeat checks cost far less because unchanged pages are not re-sent.

### The key is on the FREE TIER, with a 20-requests-per-day-per-model cap

Discovered while trying to measure a third anchor: `gemini-3.5-flash` returned 429
`RESOURCE_EXHAUSTED` on every attempt for 15 minutes. The error body named the cause:

```
quotaMetric : generativelanguage.googleapis.com/generate_content_free_tier_requests
quotaId     : GenerateRequestsPerDayPerProjectPerModel-FreeTier
quotaValue  : 20
```

**So the practical limits are:**

- 20 requests per model per DAY on the free tier. One check costs 3 marking calls
  (plus cached text batches), so roughly **six checks a day** on a given model — or
  far fewer while measuring anything.
- **Each model has its own allowance**, so a cheaper model can be used when an
  expensive one is exhausted. That is a genuine reason `model` and `visionModel`
  are separate.
- Enabling billing on the Google Cloud project removes the daily cap.

**Fixed as a result:** the app answered every 429 with *"wait a minute and press
Check now again"*, which is useless for a daily cap — no amount of waiting a minute
helps. `describeRateLimit()` now distinguishes the two and, for a daily cap, names
the model, says the limit is daily, points at billing or a cheaper model, and
reassures that nothing was lost. The vision pass also logs the real reason instead
of quietly falling back to a text-only mark, which would have looked like the app
simply getting worse. Tests: section 11n.

**Still outstanding:** the third anchor's `gemini-3.5-flash` measurement. The two
that did complete (99 -> 97, 96 -> 92) remain valid, but a third at 89 was blocked
by this daily cap. Retry on a fresh day.

### "Drag and drop does not work" — the real cause was TWO sources of truth

Reported: *"I tried to drag and drop Jacob's and it kept Rohan's portfolio on the
screen."* The drop itself **worked** — `DT - Google Slides.pdf` was saved correctly
and the endpoint behaved. The fault was in what the page DISPLAYED:

- the card took the **deck name from `data.deck`**, which describes the last
  COMPLETED scan, and
- the **PDF name from the live config**.

So after choosing a new portfolio the card still named the previous `.pptx` while
the config already pointed at the new PDF. Two sources of truth, and the stale one
won. Fixed: the live config decides, `data.deck` is only a fallback for a server
that predates the merge, and the page count is discarded when it describes a
different portfolio than the one configured.

The card also now says **"Not read yet — press Check now. Everything below still
describes the previous portfolio."** A freshly chosen portfolio has not been
scanned, so without that line the old findings look like the new portfolio's
results — which is exactly how "it kept Rohan's" was experienced.

**There was no way to remove the PowerPoint.** A `.pptx` could be chosen but never
removed, so starting clean was impossible and the old one stayed silently in use.
`Remove PowerPoint` now exists beside `Remove PDF`, and deliberately leaves the PDF
alone — the two are separate files and a PDF on its own is a valid thing to mark.

**Diagnostics added:** `saveConfig` now logs which portfolio fields moved, e.g.
`Config changed - deckFile: Rohan Ghali.pptx -> (none); visionModel: ... -> ...`.
Chasing this without that meant guessing from a file modification time. Tests:
section 11r.

### "It is unclear which portfolio is being marked" (fixed)

Reported after a drag and drop. The app said it TWICE, in two sections, in two
different ways: the upload section said `In use: <the .pptx>` and the
visual-review section said `In use: <the PDF>`. Two answers to one question reads
as a contradiction, so the student could not tell which file the mark described.

Worse, with only a PDF set the upload section said "No portfolio has been read
yet", which is simply false — the PDF *is* the portfolio at that point.

**What is actually true, now stated once at the top of the page under "Marking":**

| Set | What is marked |
|---|---|
| a PDF | the PDF — every page is looked at. Its text also comes from the `.pptx` **if one is set** |
| only a `.pptx` | that file's text; drawings and photographs are invisible to the AI |
| neither | nothing, and it says so |

The sections below now describe what each file is FOR ("PowerPoint file:",
"Pages supplied by:") rather than each claiming to be the verdict. A test asserts
the string `In use:` no longer appears anywhere, so the second verdict cannot come
back (tests: section 11p).

**Two genuine ambiguities are now flagged on that card**, because both produce a
mark nobody can explain:

- a PDF and a `.pptx` with **different names** — if they are different pieces of
  work, the mark would combine two portfolios;
- a PDF with **no `.pptx`** — the free structural checks then have nothing to read
  and the page count is unknown.

### Two reported bugs (both fixed)

**1. "Inconsistent displayed pricing".** The page and the server quoted DIFFERENT
prices for the same scan: the page said **9.22p**, the server **16.07p**, a 6.85p
disagreement. Every shared constant matched, which is why the mirroring tests
passed — only the arithmetic differed. The browser added the marking runs' tokens
but priced them at the PAGE-READING model's rate, ignoring `visionModel`. Fixed by
pricing the marking runs at the marking model's rate in
`estimateScanGbp(model, pages, opts)`; four display sites now pass the marking
model. Both sides now compute 16.07p, and a test reproduces the page's own formula
from its own constants and compares it with the server's (< 2% tolerance), so a
future divergence fails rather than shipping.

**2. Dropping a PDF warned about a mismatched name.** Reported as: dragging a
friend's PDF produced "This PDF is named ... but the portfolio in use is ...".
That was wrong twice: it blocked the legitimate case of reviewing someone else's
portfolio, and the warning described a problem that had not happened. Dropping a
PDF now CLEARS the deck in use — which removes any mismatch by construction — and
the reply states what it replaced. Tests: section 11m.

A NOTE ON CLEANUP: restoring the config after testing deleted the PDF that the
config pointed at, leaving `pdfFile` dangling and the next scan silently text-only.
Copying the file back fixed it. **Never delete a file under `state/portfolio/` that
`config.json` references** — check `pdfFile` and `deckFile` first.

### Published: github.com/rohanghali/NEA-coach

Live, public, with a working download. **The shareable link is
<https://github.com/rohanghali/NEA-coach/releases/latest>** — that always points at the
newest release, so it can be given out once and left alone.

`tools/release.sh <version>` does the whole publish: builds the package if needed, creates
the release and uploads the zip. It reads the token from the macOS keychain, where the first
`git push` stored it, and passes it to curl through a config file rather than on the command
line — a token in argv is visible to anything that can list processes.

**The README had been destroyed and this is the important lesson.** It had been replaced at
some point with the four-paragraph repository blurb written for the student, so the GitHub
page had **no install instructions, no download link and no troubleshooting** — 7 lines
where there had been 639. It was recoverable only because the full file was in the first
commit's history. Restored and brought up to date: the Releases download first, Windows
needing no install, GCSE as well as A-level, the daily allowance instead of money, and the
advice to switch to a Flash-Lite marking model **removed** — that advice was actively wrong,
because that model was measured compressing a 99/96/89 set to 81/80/66.

**A user-visible bug found while chasing a failing test.** A scan that fails, or runs
without an API key, still writes a report with `predicted: 0, maximum: 0,
weightsKnown: false` and a row for every criterion. The predicted-mark guard only checked
for rows, so it let that through and the card would have shown **"0 / 0" marks — a grade of
zero on a portfolio nobody had assessed.** The server already flagged it with
`weightsKnown`; the UI now honours it. Fixed in **v1.0.1**; v1.0.0 has the bug.

**A note on the test that surfaced it:** it asserted the stored report always contains a
percentage, which made it depend on the last scan having succeeded rather than on the
snapshot code it was meant to exercise. Rewritten to assert the real invariant — counts
always, a mark only when there was one.

### The allowance meter shows a percentage again

Reported: *"today's free allowance isn't showing as a percentage"*. That was my doing — an
earlier pass decided the percentage was dishonest while the daily limit was only the app's
assumption, so it showed a bare request count until a 429 confirmed the real limit.

**That was honest and unreadable.** "14 requests" tells a student nothing about how close
they are to running out, and the student had explicitly asked for a percentage. Corrected:

- **A percentage is always shown**, worked out from the model with the most used, so it
  reflects whichever allowance will run out first.
- **What changes is the LABEL, not whether a number appears.** Confirmed limit: exact.
  Assumed limit: marked `(estimated)`, with a footnote saying the counts are exact while
  the limit is Google's and unpublished.
- **Every model row shows its own percentage** (`14 of 20 used · 70%`), so the reader can
  see which allowance is closest to running out without doing arithmetic.
- **The bar is drawn again** whenever there is a percentage — it had been conditioned on
  the limit being confirmed, so the meter had no bar at all on a fresh key.
- The zero state shows `0%` rather than a bare `0`.

The server still distinguishes confirmed from assumed (`limitKnown`, `limitAssumed`), which
is what makes the honest labelling possible; only the client's presentation changed.

### A Past checks row could not be closed

Reported: *"you can't close the past checks expansion thing when you click on it"*. The
click handler called `open(index)` **unconditionally** — there was no toggle, so a row
opened and stayed open and the only escape was reloading the page.

Fixed by tracking which row is expanded (`openIndex`) and closing it on a second click.
Two details beyond the bare fix:

- **The open row is visually marked** (`is-open`, a tinted background and a `▾` marker).
  A state the reader cannot see is a state they cannot control: previously an open row
  looked identical to a shut one, so there was no way to know clicking again would close
  it.
- **`aria-expanded` is set** on each row, so the state is available to a screen reader
  rather than only to the eye.
- An empty detail box now collapses (`:empty { display: none }`) instead of leaving a gap
  after closing.

Tests: section 12d, which reproduces the bug directly — two clicks on the same row must
end closed — rather than only asserting the source contains a toggle.

### Past checks now name the file that was checked

Requested because the list was ambiguous: a student who has reviewed their own portfolio,
a friend's and a draft had several entries whose marks could not be told apart.

**The filename was already in the snapshot and already returned by `/api/history` — it was
simply never rendered.** So this was a display omission, not missing data, and every
existing entry shows its file immediately with no re-scan.

- A **Portfolio** column, with the filename wrapping rather than widening the table.
- An older entry written before the filename was recorded says **"not recorded"** rather
  than rendering an empty cell, which would read as a fault.
- The **marking model** is shown under the filename. This one matters more than it looks:
  the same portfolio scores differently across models, and without it a mark that moved
  because the model changed looks like a mark that moved because the work changed.
- Opening an entry repeats the portfolio name, so a clicked row is unambiguous.

Tests: section 12c.

### Tidy-up pass: first impression, summary, and the prose

Asked for after the student looked at it on their own machine: the ring was the first
thing they saw and it was **red**, the summary beside it read like documentation, and the
page explained the app's own machinery instead of the student's work.

**1. The ring was alarming, and the cause was two different numbers sharing a label.**
The ring fills in proportion to pages free of high-severity findings, but the label beside
it read "29 points to check" — those are not the same measure, and the band keyed the
COLOUR off the ring. On the real portfolio that produced `55% -> ring-high -> --high`, and
`ring-critical` is `#b42318`, a red alert on a perfectly normal portfolio.

Now the lowest two bands use `--medium` (amber), so red is reserved for the individual
findings where it actually means something, and the label reads **"points worth a look"**
rather than implying a status.

**2. The summary was rewritten in plain language.** It used to open "These are points to
check, not a mark" and talk about "mechanical checks" and "presentation and completeness
points" — the app explaining itself. It now answers the three questions a student has, in
order: how much, which one matters, where to start.

> Found 29 things, and 24 of them cost you marks. Start with number 1 — it is the one costing you most.

**3. Prose that explained the machinery was cut**, because a student should not have to
read how the checks work to find out whether their coursework is in trouble:

| Cut | Kept |
|---|---|
| "mechanical checks", "presentation and completeness points" | how many, and which matter |
| 4 paragraphs on what deep thinking does internally | what it helps with, and that it does not change the mark |
| the reasoning-token accounting | it is slower and uses more allowance |
| "pages are fingerprinted… several pages travel in one request" | the line above the button already says the request count |
| five lines of caveats under the predicted mark | two: how it was judged, and that it is an estimate |

The Information panel, Statistics and Slides are all **collapsed by default**, so a
first-time visitor sees only what they must: which portfolio is being marked, the summary,
the Check button, and then the findings.

**A self-inflicted bug worth noting:** my first rewrite produced *"24 of them worth
markss"*, because `plural()` appends an "s" and I handed it a whole phrase. Caught by
printing the assembled sentence rather than reading the diff — which is the only way that
class of error shows up.

### Sharing it: packaged for Windows, and on GitHub

Written up for the student in **`PUBLISHING.md`** — account, repository, push, build,
release, and the link to send people.

**`tools/package.sh`** builds the distributable. It exists because the working folder
holds **556 MB of the student's own work** in `state/` — the report, a copy of the
portfolio, the scan history and the API key — plus a 104 MB Node runtime that only runs
on the machine that made it. Neither belongs in something sent to someone else.

| Build | Contains | Size |
|---|---|---|
| Windows | app + `node.exe` (79 MB, uncompressed) | **31 MB zipped** |
| Mac | app only; the launcher downloads the matching Node | ~1 MB |

The Windows zip needs **no installation at all** — Node is inside it, and the `.bat`
already looks for `runtime\node.exe` before PATH or the standard install folders. It
also carries a `START HERE.txt` for someone who has just unzipped it.

**Two things found while doing this, both of which would have bitten a recipient rather
than us:**

1. **`.gitignore` had to exist before anything was committed.** `state/` contains the API
   key in plain text, and a key pushed to a public repository is scraped within minutes.
   The ignore file also covers the Dropbox `conflicted copy` files, which have contained
   a key before. **Verified by initialising the repository and inspecting every staged
   path: 28 files, zero from `state/`, `runtime/` or `dist/`, and a grep for the key's
   prefix across all 28 found nothing.**
2. **The `.bat` had Unix line endings** — 172 bare LFs and no CRLF, because editors on
   macOS strip them. `cmd.exe` can mis-parse that, and it would fail on the recipient's
   machine. `package.sh` now converts the file at build time with `perl -pi`, and the
   packaged copy is verified as CRLF.

**Monetisation, answered honestly:** GitHub gives a free permanent link and **cannot take
payment**. If the student charges later, that is Stripe or Gumroad. More importantly,
**each buyer must supply their own Google API key** — Google's terms do not permit
reselling API access, and a shared key would make the seller liable for every check every
buyer ran. Both points are in `PUBLISHING.md` so the decision is not made blind.

Committed locally as a fresh repository (`4bdfc00`), ready to push once the student gives
their GitHub username.

### The predicted-mark panel was restructured (student: "looks a bit messy")

It was a flex row: a **128px column** holding the percentage, the marks, the run count,
the caution and the calibration note, with the explanation beside it. Three percentages
and three separate notes in a 128px column is what looked wrong — the figure wrapped and
the notes ran together. The calibration note was also a `<div>` inside a `<p>`-styled
flex item.

Now: **the number is one block on the left, everything supporting it stacks in a column
on the right**, each line with room and an order to follow. On a phone it becomes one
column, figure first.

**A second fault surfaced while fixing it, and it mattered more than the layout.** With a
correction applied the panel said **"about 80–84%"** as the headline while the app
believed **89** — two answers on screen with nothing saying which governed. That is part
of why it read as a mess, and it was a real contradiction rather than a cosmetic one.

So when a correction applies:

- the **adjusted figure is the headline** (`about 89 of 100 marks`), because it is the
  app's best estimate and the raw one is a known under-read;
- the AI's own reading is **not hidden** — it moves into the supporting column as a
  quiet line (`The AI's own reading was about 80–84%`), because the raw figure is the
  evidence and hiding it would make the adjustment look like a fudge;
- the reason sits in a highlighted box, introduced by **"Why this is higher than the AI
  said"**, rather than repeating the raw figure back (the note already opens with it —
  the first attempt duplicated the sentence verbatim).

Tests updated to pin the structure, so it cannot quietly revert to a cramped row.

### SCORING: corrected by what the app has been MEASURED to get wrong

The student asked for this to be acted on, and gave the missing data point: Jacob's
portfolio (real **89**) scored **80-84%** on the default marking model, not the 66 the
cheap model gave it. That made a **third anchor for `gemini-3.5-flash`**, and the first
below 96 — the only thing that could test whether a fitted line was an artefact of two
points. It was not:

```
real = app x 0.722 + 29.2      max residual 0.4 marks
  app 97 -> 99   app 92 -> 96   app 83 -> 89
```

**What was wrong:** the calibration tool existed, was documented, and was **never
applied**. `grep -n "calibration" server.js lib/ai.js public/index.html` returned
nothing — no code at prediction time read the anchors at all. The app reported the raw
model figure, which reads **2-6 marks LOW**.

**What now happens** (`lib/calibration.js`):

- Anchors are **grouped by marking model**, because the two models fail differently:
  `gemini-3.5-flash` reads 2-6 low, while `gemini-3.1-flash-lite` **compresses** (it
  gave a 99, a 96 and an 89 portfolio 81, 80 and 66). Both sets are in the file, and
  the cheap model corrects from 66 -> 89. Pooling them once produced the nonsense fit
  `real = app x 0.043 + 93.7`.
- The correction is applied **only inside the measured range (83-97)**. Below it
  nothing has been measured, so the raw figure is left alone and the app says so.
  **Extrapolating a fudge factor into unmeasured territory would make the app
  confidently wrong in a new way** — worse than being visibly uncertain. A weak
  portfolio remains the one real unknown.
- The **raw AI figure stays visible** next to the adjusted one, with the reason in
  words. A predicted mark shown without its error bar invites more trust than it has
  earned, and an unexplained adjustment is a fudge. The UI and the export both carry
  the note.

**Measured after the change:** the default model lands within **1 mark on all three
anchors** (0, 0, 0); the cheap model within **1** ( -1, +1, 0). Jacob's report now
reads **83 raw -> 89 corrected**, against a real 89.

**One test lesson:** I first asserted the two slopes differed by >0.15. The real
difference was 0.13, so it failed — and the threshold was arbitrary anyway. Replaced
with the meaningful claim: applying the *wrong* model's correction must be visibly
worse, which is what grouping actually protects.

### Four more requests: money out, honest meter, stats, export

**1. Money removed entirely.** The key is free-tier, so a price is always zero. Removed:
the per-model `estimatedScanGbp` from the API payload, the price-based model sort, the
whole "switch to a cheaper model" suggestion (there is no saving to make), the dead
`scanNow` computation in Settings, and `moneyPrecise`/`money`. **No currency symbol
remains in code the browser runs** — a test strips comments first and asserts this,
because a comment explaining the removal legitimately contains "£0.00".

**2. The percentage was not honest, and is now only shown when it can be.** It came
from `ASSUMED_DAILY_REQUESTS = 20`, a figure **Google never confirmed** — the
rate-limits page deliberately does not publish per-model daily numbers. A confident
"70% of your allowance" derived from a guess is worse than saying nothing. Now:

- `percent` is **null unless a 429 has told us the real limit**, and the UI shows the
  exact **request count** instead, which is the one thing the app genuinely knows;
- the `Math.min(100, …)` clamp is gone, so exceeding the limit reads as over 100
  rather than as exactly 100;
- the pre-scan hint no longer says "6 of 20 left" from an assumed 20.
- `worstPercent` is null when no limit is known, rather than a fabricated worst case.

**3. Stats card.** Computed on the SERVER (`buildStats`) because the full text of each
page never reaches the browser — only a summary per slide, and shipping every page's
text on every poll would be megabytes. Covers words, pages, images, notes, average
words per page, **typing time at 40 wpm** (the rate is named so the figure can be
judged), **marker reading time at 150 wpm**, words per image, longest sentence, pages
with no words, heading coverage, the words the portfolio leans on, and evidence balance
per criterion from the pages the AI cited. All measured; nothing is a mark.

**4. Export.** `GET /api/export` returns **one self-contained HTML file** — no images,
no external stylesheet, no fonts, so it opens anywhere and prints to PDF or Word. It
carries the predicted mark, per-criterion table, every finding grouped by tier with
evidence/why/what-to-do, a **printable tick-list**, strengths, suggested additions, the
document stats, and how the report was produced. Built server-side because the full
finding text does not all reach the client. Every string is escaped — asserted with a
deliberate `<script>` and `<img onerror>` payload, because the report contains student
writing and an unescaped angle bracket would silently break the document.

**5. "What costs your allowance" became the Information panel**, a collapsible card at
the bottom, and gained four more topics: what the app does and cannot do, where your
work and key are kept (and to delete `state/` before sharing), changing the API key,
and why a PDF skips some checks.

NOT DONE, on the student's instruction: the AI chat, parked for a later pass.

### EVERY AI CHECK WAS FAILING — `onRequest is not defined`

The student reported: *"the last check stopped early: requestsBefore is not defined
... if it's talking about rate limits I have loads left."* Two separate faults, one of
them severe.

**Fault 1 — `requestsBefore is not defined`.** The quota baseline was declared inside
the AI-only branch but read outside it, so any scan that skipped the AI (PDF-only, no
API key) died after the local checks. Fixed by declaring it at the top of the try
block.

**Fault 2 — every AI request threw `ReferenceError: onRequest is not defined`.**
This was the serious one and it was invisible in review. The counting code added
inside `callWithRetry` referred to `onRequest` — but `onRequest` is destructured in
`callGemini`, **not** in `callWithRetry`. So every request threw; the surrounding
`catch` treated it as a network failure and reported *"The AI request for pages 1-3
failed: onRequest is not defined"*.

**The real AI checks were therefore broken, not rate-limited.** That is why the
student saw "stopped early" while Google showed allowance remaining — Google was never
successfully asked. Found by dumping `result.ok`/`result.error` from a test run rather
than reasoning about it, after several wrong guesses.

**Fault 3 — the meter over-counted.** It counted every request it SENT, including
429/5xx refusals (which consume nothing) and `/v1beta/models` GETs (which are not
`generateContent` calls and do not touch
`GenerateRequestsPerDayPerProjectPerModel`). It reached **27 against a limit of 20**
while Google was still serving checks. Now:

- only a **served (HTTP 200)** response is counted, and
- only **once per logical request**, so a retry that eventually succeeds is not counted
  twice, and
- `listModels` and the key test are **not** counted at all.

Live proof after the fix: a sample-deck pass returned `ok: true`, the counter fired
**2** times, `batchesRun` was **3** (2 batches + 1 cross-check), and a real
`noteRequest` moved the stored count 12 -> 14.

**A wrong note I had to discard:** the earlier claim in this file that "one full
60-page check is about 9 requests" was based on `batchesRun`-style counting including
the cross-check. The request count per check is what the counter now records exactly,
so trust that number over the estimate. Tests: sections 11x and 11y.

### The costs section now tells the free-tier truth

"What costs money, and what does not" was written when the student was paying per
check: it said the AI review *"is billed by Google to your account"*. That became
false the moment the key was deliberately put on the free tier.

Rewritten as **"What costs your allowance, and what does not"**, and — more
importantly — **the rest of the page was changed to match it**, because prose alone
would have left the UI contradicting itself:

| Was | Now |
|---|---|
| model dropdown: "about 6.2p per full scan" | "about 9 requests per check (about 2 checks a day)" |
| Settings model line: a price per scan | requests per check |
| pre-scan preview: money on each model | requests, and how many checks fit today |
| history entry: "£0.06 spent" | "9 requests used" |
| "cheaper model" nudge: prices and a saving factor | a second key/model has its **own** allowance |

**Rule adopted: no money figure is shown for a free key.** It reads £0.00 forever and
says nothing about the thing that actually runs out. `moneyPrecise()` is kept and
documented for the paid-key case only.

The section now states the real numbers (**~9 requests per full check**, **~20 per
model per day**, so **~2 checks a day**), explains that the structural checks never
touch the allowance, repeats the PDF limitation where costs are explained, and lists
the three escape routes in order of speed: **paste another key**, wait for midnight
Pacific, or use a cheaper model / enable billing.

`snapshotOf()` gained `requestsTotal` so the history can say what a check spent now
that it no longer reports money. Tests: section 11w.

**Still outstanding, deliberately not touched:** `estimateScanCostGbp` and
`estimateScanGbp` still compute money, and the server still sends `estimatedScanGbp`
per model in the suggestion payload. Nothing displays those any more, so it is dead
weight rather than a bug — worth removing in a dedicated pass, not as a drive-by
during a wording change.

### Changing the API key from the interface

The app used to say the key was **"never shown or typed here"** — so swapping it meant
hand-editing `config.json`. But the free tier runs out, so swapping is routine, not
setup. Settings now has a **Paste a key** box with:

- **Test key** — validates the key by calling `listModels`, a plain GET. This is
  deliberate: **testing must not spend generation quota**, or it would be unusable at
  exactly the moment it is needed, because the allowance is spent. Confirmed on the
  real exhausted key — it reports `ok: true`, 45 models, and names the three the app
  uses. That is the correct answer: the key is valid, its allowance is simply gone.
- **Save key** — stores it immediately, no restart, no scan needed.
- **Forget key** — clears it.
- The panel says which key is loaded **by its last four characters**, because two keys
  look identical until one runs out.

**The key is never sent back to the browser.** `publicConfig` exposes `keyTail` (last
4 chars) and nothing else; a test asserts no part of the key appears in the public
config.

**The quota message now points here first:** when the daily limit is hit, the app says
to paste a different key in Settings — each key has its own daily allowance — before
suggesting billing or waiting. That is the fastest fix and now a two-click one.

A near-miss worth recording: the first version of the UI used `--surface`, `--good`
and `--bad`, two of which do not exist in the stylesheet. **Unstyled CSS fails
silently** — no error, just an unstyled box — so the check that catches it is the
`var(--x)` vs declared-tokens diff, which now reports zero.

Tests: section 11v.

### "The marking tab doesn't change" — a field missing from the merge

Reported twice, with two different causes behind the same symptom.

**First time:** the card read the DECK name from `data.deck`, which describes the
last COMPLETED scan, so after choosing a new portfolio it still named the old one
until a scan ran. Fixed by preferring the live config.

**Second time (the real one):** the card read the PDF name from `config.pdfName`,
and **`publicConfig` returned `pdfFile` but never `pdfName`.** `buildStateResponse`
does `Object.assign({}, report.config, publicCfg)` — the live config wins for keys
it HAS, but a key it does not have **survives from the report**. So the report's
`pdfName`, written whenever that report was generated, was shown forever, while
`pdfFile` beside it was live. The card could display `Rohan Ghali.pdf` while the app
was marking `DT - Google Slides.pdf`.

**Fixed at both ends, deliberately:**

1. `publicConfig` now derives `pdfName` from the live `pdfFile`, so the two can
   never disagree — and the key exists, so the merge overwrites the stale one.
2. The card also derives the name from the path itself and prefers the path when the
   reported name disagrees, so it is correct even against an older server.

**The general lesson, worth remembering:** `buildStateResponse` merges the live
config OVER the report's copy, so ANY display field the client needs must be present
in `publicConfig`. A missing key does not become undefined — it silently keeps the
value from the last scan. Tests: section 11u.

### Structural checks now run on a PDF too

A PDF-only student previously got NO structural checks: `findDeck` required a
`.pptx` and errored without one, so dropping a PDF and pressing Check failed
outright. A PDF export is a completely normal way to hand in work, so this was a
real gap.

**How it works.** Two new files, both zero-dependency:

- **`lib/pdftext.js`** — a pure-JavaScript PDF text extractor (Node built-ins only:
  `fs`, `zlib`, `Buffer`). No shelling out, no `osascript`, no `pdftotext`, so it
  works identically on macOS and Windows. Resolves the page tree, decodes
  FlateDecode streams (with PNG/TIFF predictors), walks the text operators
  (`Tj`/`TJ`/`'`/`"`, `Td`/`TD`/`T*`/`TL`) and maps glyph codes through the fonts'
  `/ToUnicode` CMaps. Robustness: junk, empty, truncated and random buffers all
  return a result with a warning rather than throwing.
- **`lib/pdfdeck.js`** — turns that into the same shape `lib/pptx.js` produces, so
  nothing downstream cares which file it came from. Adds a content fingerprint per
  page (absent otherwise, which would make every page look new on every scan).

**Accuracy check:** extracted character counts match macOS's own Quartz extractor to
within **0.02%** (Jacob, 71,341 vs 71,329) and **0.6%** (Rohan, 142,629 vs
141,749). Not a single empty page on either file.

**The honest limitation.** A PDF carries no speaker notes, and it has no text
boxes. Rather than feed the rules invented numbers, each rule is TAGGED with what it
needs (`rule(name, fn, requires)` — `text`, `images`, `geometry`, `fonts`, `notes`)
and the ones a PDF cannot support are **skipped and named**:

| Source | Rules run | Skipped |
|---|---|---|
| `.pptx` | 18 of 18 | none |
| `.pdf` | 13 of 18 | page-too-busy, uncaptioned-image, missing-notes, text-overflow, off-canvas |

**The PDF extractor now reads real geometry** (`lib/pdftext.js`). Every shown string
is emitted as a positioned run carrying the text matrix, the font size and the
BaseFont name — values that were being parsed and then thrown away. `lib/pdfdeck.js`
turns those runs into per-line text blocks, so `text-too-small` and
`font-inconsistency` now run on a PDF with **exact** inputs. Measured on Jacob's
44-page PDF: 13 of 18 rules, `text-too-small` firing on 34 pages, and
`font-inconsistency` reporting "12 different heading sizes".

**Why the other geometry rules still do not run.** Line WIDTH is an estimate, and
these are the documents where that bites: a Google Slides export uses Type0 CID
fonts and carries **no `/Widths` array at all**, so per-glyph advances are not
available without parsing the embedded font program. Measured against the real
portfolio the estimate over-read a line by about **13%**, which turned **106
correctly placed lines into "runs off the page" findings**. So the width-dependent
rules (`text-overflow`, `off-canvas`) stay skipped rather than crying wolf about a
page the student can see is fine. `page-too-busy` needs geometry only for its
evidence string and is skipped for the same reason. A rotation guard also exists:
`deck.geometryOk` is false if any page carries a non-zero `/Rotate`, because a
rotated page's coordinates are in a different frame from the rendered page.

A rule that never runs finds nothing, and "nothing" reads as a **pass** — which is
why they are listed in the report and explained in the UI (which names the categories
dynamically, so the message updated itself when font sizes started running). Real
findings do come out: 8 on Jacob's PDF, including a near-empty page and a spelling
inconsistency, matching what the same file shows as a `.pptx`.

**Cross-checked against PowerPoint, deliberately.** The same portfolio as `.pptx`
gives 12 `text-too-small` findings; as a PDF it gives 13. That is agreement, not a
regression — and it is the test to repeat if these rules are ever retuned.

**Three bugs found and fixed on the way:**

1. **`requestsBefore is not defined`** — a crash I introduced with the quota
   counting: the variable was declared inside the AI-only branch but read outside
   it, so EVERY scan that skipped the AI (PDF-only, or no API key) died after the
   local checks. Moved to the top of the try block.
2. **`Cannot read properties of undefined (reading 'toFixed')`** — the report assumed
   `deck.slideWidthIn`/`slideHeightIn`/`aspect` were numbers, which `lib/pptx.js`
   supplies and a PDF does not. `pdfdeck` now provides real dimensions from the PDF,
   and the report is null-safe anyway.
3. The page tree, xref and `/ToUnicode` stream resolution in `pdftext.js` (found by
   the subagent that built it, against the real files).

Tests: section 11t.

### Renamed to NEA Coach

The app was called DTA Coach, which read as the SUBJECT (Design and Technology)
rather than the tool. Everything the user sees now says **NEA Coach** — the page
title, the launcher banners, the docs — and the app folder is `NEA-Coach`.

123 occurrences in 19 files, replaced in this order so nothing was double-replaced:
`DTA_COACH` → `NEA_COACH` (the env var), `dtaCoach` → `neaCoach` (the localStorage
keys), `DTA-Coach` → `NEA-Coach` (the folder name), `DTA Coach` → `NEA Coach`.

**`DTA` still appears where it means the qualification**, which is correct and must
stay: `spec/README.md` says "Adding the real Eduqas DTA specification".

**Two traps hit while doing this, both worth knowing:**

1. **A blanket rename broke an absolute path.** `state/config.json` held
   `deckDir: /Users/.../DTA-Coach/state/portfolio`, and the replacement turned the
   PARENT directory into `NEA Coach` too, producing a path that did not exist.
   `deckDir` is now **relative** (`state/portfolio`) and resolved against the app
   folder, so it survives any future move. `pdfFile` must stay absolute (it points
   at a file, not the app), so it was repointed by hand after the folder rename.
2. **The app name wrapped across a line break** in two documents, so
   `Start DTA
Coach.bat` slipped past a plain string match. Found by an
   exhaustive case-insensitive grep, not by the replacement count.

**The launcher files were renamed too:** `Start NEA Coach.command` and
`Start NEA Coach.bat`. The env var they pass is now `NEA_COACH_PORT` — setting the
old `DTA_COACH_PORT` is silently ignored, which briefly made a test server start on
4173 instead of the port intended.

**Not renamed: the parent folder** `/Documents/APPS/DTA Coach`. It contains this
project's folder and two portfolio PDFs. Renaming it is the user's call, and would
mean reopening the app from the new path.

### What is now DONE, and what is merely UNVERIFIED

Done and measured:

1. **Comparing two scans.** Exact matching for the free checks (same rule, same
   page), per-criterion mark movement, and a headline split into substance versus
   mechanics. AI findings are deliberately NOT diffed — the model rewords and
   re-groups them, which once produced "21 fixed, 22 new" on an untouched file.
2. **Dark / light mode**, following the system by default with an explicit
   override, applied before first paint.
3. **"Deep thinking" explaining itself** — what it adds, when it helps, when it
   does not, and what it costs.
4. **Predicted-mark accuracy.** The cause was the MARKING MODEL, not the prompt.
   Flash-Lite was 16-26 marks low with a ~20-mark spread; `gemini-3.5-flash` is
   within 2-4 marks on both anchors with a 2-mark spread, and is now the default.

Not yet verified, and should not be claimed:

- **Accuracy is only proven at the top of the range.** Both anchors are strong
  portfolios the app now marks 92 and 97. Nobody has measured a WEAK portfolio, so
  it is unknown whether the app would over-mark one. The bundled sample deck is a
  .pptx and there is no renderer, so it cannot be used as an anchor without a PDF.
- **The phone layout has never been viewed on a phone** since the recent additions.
  Test 11i proves the mobile rules exist; it cannot prove they look right.
- **The Windows launcher has never run on Windows.**
- **`tools/calibrate.js` has only ever been pointed at these two PDFs.**

### Known limitations, stated honestly

- **Live AI calls untested.** Parsing, model selection, batching, caching and cost
  maths are unit-tested; the network leg is not.
- **The launcher has never run on Rohan's iMac.** It detects architecture
  (`arm64`/`x86_64`) and downloads the matching Node. It is unverified on that
  machine.
- **The Windows launcher has never run on a real Windows machine.** `Start NEA
  Coach.bat` looks for `runtime\node.exe`, then `node` on PATH, then the standard
  install folders (`%ProgramFiles%\nodejs`, `%LOCALAPPDATA%\Programs\nodejs`). If
  no suitable Node is found it prints instructions to install the LTS from
  nodejs.org; it deliberately does **not** download a binary, because a 45 MB
  download inside a `.bat` is unreliable and antivirus-flagged. The file has been
  read through carefully for quoting faults, and the Node one-liners it uses
  (the 18+ version gate and the `/api/health` probe) were run and checked, but
  the batch syntax itself was never executed, because no Windows machine was
  available. The parts most likely to need a correction are the `%~dp0` quoting
  and the `if`/`goto` flow. It is saved with CRLF line endings.
- **PDFs in `spec/` are not read.** The app reports a skipped PDF rather than
  silently ignoring it. The Eduqas spec was converted with `pdftotext`.

---

## 8. How to ask an AI to change this

Any capable AI can edit this project, including on a free tier, because there is
no build step and no dependencies. What it needs:

1. The files. The whole app is about 3,000 lines. `server.js` (~900),
   `lib/*.js` (~2,400), and `public/index.html` (~2,900) are the bulk.
2. This file, so it knows the rules in section 3.
3. **Tell it to test against both decks** using section 4.

A good request looks like:

> Read `NEA-Coach/CONTINUE-HERE.md` first. Then change the caption check so it
> also accepts a caption to the left of the image. Test against both
> `sample/AQA-GCSE-Portfolio-SAMPLE.pptx` and the real portfolio, and confirm the
> planted flaws on pages 6 and 7 are still caught while page 14 is still excluded.

Record what you want in **`WISHLIST.txt`** in this folder as you think of it. That
way nothing gets lost, and any future session can read it straight away.

---

## 9. A note on the numbers in this project

Every figure quoted anywhere in this project came from a measurement, not a guess.
If you change a threshold or a rule, **re-measure and update the README**. The
earlier low-resolution rule was set to 150 DPI by intuition and produced dozens of
false complaints about a 99/100 portfolio; the distribution was 200 images above
300 DPI, 100 at 150-300, and 39 below 150. Intuition was wrong; measurement fixed it.
