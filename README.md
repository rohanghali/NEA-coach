# NEA Coach

A read-only portfolio reviewer for a Design & Technology NEA portfolio, at **GCSE**
(AQA 8552) or **A-level** (Eduqas). It runs entirely on your own computer — Mac or
Windows. Give it your PowerPoint or a PDF export; it tells you what would raise your
mark and in what order to fix it, and **never edits your file**.

**It is free.** You supply your own Google API key, which needs no card and no
account, and the app tracks the free daily allowance rather than money.

---

## The guarantee: your PowerPoint is never touched

This matters, because portfolio regulations require the work to be yours.

- The `.pptx` is opened **read-only**. The code that opens it has no write path at all.
- The app never launches PowerPoint. It reads the file format directly.
- Nothing is uploaded anywhere except the **text of your slides**, and only if you
  have added an AI key. See "What leaves your computer" below.
- The app creates no files inside its own folder except a downloaded `runtime/`
  and its `spec/`. Everything about you — your settings, your key, the saved report
  and the copy of your portfolio — lives in **your own user data folder**, outside
  the app, so unzipping a newer version never loses any of it. See
  "Where your data lives" below.

You can verify this yourself: make a copy of your portfolio, note its size and
modified date, run a scan, and check that both are unchanged.

---

## Quick start

**Download it from the [Releases](../../releases/latest) page** and unzip it.

If this repository is ever made **private**, that page and every release link stop working for
anyone without access — a private repository returns 404 for release files even when a token is
supplied, because they are only served through the API. `tools/get-release.sh` handles both
cases and is the reliable way to fetch a build.

1. **Start it.**

   **On Windows:** double-click `Start NEA Coach.bat`. **Nothing needs installing** —
   the copy of Node.js the app runs on is already inside the folder you downloaded.

   If Windows shows a blue "Windows protected your PC" box, click **More info** →
   **Run anyway**. That is SmartScreen being cautious about a new, unsigned program,
   not a warning that anything is wrong.

   **On a Mac:** double-click `Start NEA Coach.command`. The first time, it downloads
   the JavaScript engine it needs (about 45 MB, one time only).

   If macOS refuses to open it, **right-click it and choose Open**, then confirm.
   That is the normal macOS prompt for any downloaded file.

   **In both cases, leave the black window open while you use the app.** The app runs
   inside it, so closing the window stops it.

   *Running from source instead?* You need Node.js 18 or newer
   (<https://nodejs.org>, take the **LTS** installer), then run `node server.js`.

2. **Choose your portfolio.**
   Open **Settings**, and under "Get started — choose your portfolio" pick your
   `.pptx` file. The app copies it into the `portfolio` folder inside your data
   folder (see "Where your data lives") and reads that copy, so the file you pick is
   never changed, moved or renamed. You can
   also type or paste a folder path further down in Settings instead, if you would
   rather keep the file where it is.

3. **Press "Check now".**
   The free checks run in about a second. You get a ranked list of what to fix first.

4. **(Optional) Add an AI key** for the judgement-based review, which is the part
   that reads your work against the mark scheme. See `SETUP-AI.md`.

---

## What it checks

### Free checks — instant, offline, no cost

| Check | Why it matters |
|---|---|
| Missing page titles | A moderator cannot tell what a page is for |
| Overloaded pages (400+ words) | Communication marks are lost in walls of text |
| Unbroken paragraphs (120+ words) | Gets skipped, so the evidence stops counting |
| Text too small (under 10pt) | Unreadable when printed or projected |
| Genuinely blocky images (under 72 DPI) | Only flagged when actually unreadable — see the calibration note below |
| Images shrunk too small to read | Plenty of pixels but placed under 1.2in wide |
| Images with no caption | An unannotated image earns almost nothing |
| Near-empty pages | Wastes a page of a long submission |
| Duplicate pages | Reads as an accident, not development |
| Text overflowing its box | Adapted to respect PowerPoint's autofit settings |
| Elements off the printable page | Disappears when exported to PDF |
| Text auto-shrunk to fit | Looks fine on screen, unreadable on paper |
| Inconsistent heading sizes | The cheapest way to make a document feel designed |
| British/American spelling clashes | "aluminium" vs "aluminum", "colour" vs "color" |
| Common misspellings | Costs nothing to fix, quietly costs credibility |

### IMPORTANT: the free checks are not a mark

The app does **not** show a mark out of 100 based on these checks, and it must
never be read as one. These rules judge *presentation and completeness*. They
cannot judge the quality of design thinking, which is what actually earns marks.
A portfolio can be visually untidy and still score very highly.

An earlier version did show a number out of 100 in the ring. On a real 60-page
portfolio that was awarded **99/100 at GCSE**, that number read 51. It was
misleading and it undermined trust in every other finding, so it was removed.
The ring now shows a plain **count of points to check**, and a mark estimate
appears only once the AI has assessed the work against the real mark scheme.

### Calibration against marked work

The image checks were originally set for print perfectionism — anything under
150 DPI was flagged as a defect. Run against the same 99/100 portfolio that
produced a wall of complaints:

| Effective DPI | Images |
|---|---|
| 300+ | 200 |
| 150–299 | 100 |
| 100–149 | 29 |
| 72–99 | 10 |
| under 72 | 1 |

Sub-150 DPI images are **normal and acceptable** — usually web-sourced reference
photographs or photographed sketchbook pages. Flagging them buried the real
problems. The threshold is now 72 DPI (genuinely blocky) at low severity, and the
far more useful opposite check was added: a high-resolution image shrunk to a
postage stamp is unreadable for a completely different reason and was previously
being ignored.

Caption detection was also wrong twice. It first skipped any page with more than
60 words, which made it inert on every real portfolio page. It then looked only
*below* each image, but this portfolio captions images *beside* them, so it
reported 41 pages at once. It now assesses each image individually within a band
around it, and correctly distinguishes "2 of 5 images lack a caption" from
"none of the images have captions".
| Pages with no visual evidence | It is a design portfolio, not an essay |

### AI review — needs a free key, uses part of the daily allowance

- **Looks at your pages, not just your words.** Add a PDF export of your portfolio
  and the AI sees every sketch, photograph, CAD render and annotation. This matters
  a great deal: a real portfolio that scored 99/100 was predicted at **75%
  text-only** and **88% once the pages could be seen**, and the criterion it got
  most wrong was the most visual one. Without a PDF the review falls back to
  reading the text, and says so on screen.
- Reads each page against **your mark scheme**, not generic advice.
- Separates description from evaluation, which is where A-level marks are won.
- Builds a **coverage grid**: which criteria your portfolio evidences, how strongly,
  and on which pages, judged across the whole portfolio rather than 12 pages at a time.
- Estimates a **predicted mark for each criterion and a percentage**, so you can
  see which part of the portfolio is costing you most. The pages are reviewed
  **three times and the median taken**, so a single odd answer cannot decide the
  number. This is an estimate, not a grade: your teacher marks the real thing and
  the moderator can move it.

  **The model matters more than anything else here, and this is measured.** Against
  three real portfolios whose marks were known (99, 96 and 89), the cheap model
  predicted 81, 80 and 66: not merely low, but **compressed** — it squeezed a
  10-mark spread of real quality into a 15-mark spread at the wrong end. A Flash
  model predicted 97, 92 and 83, which reads 2 to 6 marks low but keeps the order
  right. So marking uses the stronger model and the page-reading model stays cheap.

  **The predicted mark is then corrected against those three known portfolios**, and
  the app says so on screen: it shows the raw figure, the adjusted one and why they
  differ. The correction is only applied when the raw figure falls inside the range
  the app has actually been measured on (83–97). Below that nothing has been checked,
  so the raw number is left alone and labelled as a guess rather than quietly adjusted.
  A weak portfolio is the one case this app has no evidence for, and it says as much.
- Checks every "Worth adding" suggestion against an outline of the whole
  portfolio before showing it, so it never asks for something on another page.
- Ranks findings so the highest mark-costing problem is always first.
- **Compares each check with the last one**, showing what you fixed, what is new,
  which criterion moved, and whether the predicted mark went up or down.

### Which portfolio is being marked

The card at the top of the page always says exactly what the next check will mark,
because it used to say it in two places in two different ways and that was
genuinely confusing:

- **A PDF is marked** — every page is looked at. If a `.pptx` is also set, the
  written text is taken from it as well. This is the accurate combination.
- **Only a `.pptx`** — the AI reads the text, so sketches, photographs and CAD are
  invisible to it and the predicted mark reads low.

Two situations are flagged on that card, because in both the mark would be
impossible to explain: a PDF and a `.pptx` whose **names differ** (they might be
two different pieces of work), and a **PDF with no `.pptx`** (the free structural
checks then have nothing to read).

### Spelling and wording

A dedicated set of spelling checks, deliberately kept to **clear-cut errors only**:
about 40 of the most common misspellings (*recieved*, *seperate*, *definately*,
*enviroment*, *untill*, *acheive* …). Each one quotes the wrong word, the right
word and the pages it is on, so it is a two-minute job to clear.

**Words that are correct are never flagged**, even where they look wrong:
*millimetres*, *aluminium* and *medieval* are all fine, and *dependant* is correct
British English when it means a person. A spell list that flags a student's own
subject vocabulary teaches them to ignore it, so anything arguable was removed.

Words that are real but easily swapped (*affect*/*effect*, *loose*/*lose*) are
deliberately **not** flagged. Both words are correct, so only the writer knows
which was meant, and a rule that guesses would be wrong often enough to be worse
than nothing.

### Seeing what changed on each page

The **Slides** table has a **Changed** column saying when the app last saw each
page's content change — "Edited 2 days ago", "First seen this scan", "Moved only",
or "Unchanged".

PowerPoint does not record this: saving a `.pptx` rewrites the whole archive, so the
file timestamps and revision counters cannot tell you which *page* changed. The app
keeps its own record instead, and compares pages by **content** rather than by
position — so inserting a page near the front marks that one page as new and leaves
the rest alone, instead of reporting the entire deck as edited.

### Past checks

A **Past checks** list of your last 25 scans: when each ran, how many pages changed,
the predicted mark and the finding count. Select one to see the most important
things it flagged. Each entry is a small summary rather than a whole report, so the
history stays fast and never grows without limit.

### Drag and drop

Drop a file anywhere on the page. A `.pptx` becomes your portfolio; a PDF becomes
the visual review. Anything else is refused by name, so a dropped `.ppt` or `.zip`
says so instead of appearing to do nothing.

### Comparing two checks

From the second check onwards, a **"Since your last check"** panel shows the
predicted mark before and after, how each criterion moved, the problems that are
now fixed, and anything new. The previous report is kept in
`<your data folder>/previous-report.json`; only one is kept, because "since last time" is the
question people actually ask, and a full history would grow without bound.

Findings are matched by **what they are about** — the rule and page for a free
check, the title for an AI finding — not by their internal reference number. That
matters: the reference number increments on every scan, so comparing it directly
would report every finding as simultaneously fixed and new.

### Adding the PDF (this is the important step)
In PowerPoint: **File → Export → Create PDF/XPS**, then add it in
**Settings → Visual review**. The app stores it next to your deck and sends it to
Google as part of a scan. Nothing else changes: the `.pptx` is still opened
read-only, and the PDF is only ever read.

- The PDF must be under about **18 MB**. If it is larger, export with "Minimum
  size" compression or export a page range.
- A scan uses **more of the daily allowance** once a PDF is supplied, because the page
  images travel as well as the text — roughly double a text-only check.
- **Dropping a PDF makes it the portfolio being marked.** That is what dragging a document
  onto the page means, so the app says plainly what it replaced rather than leaving a stale
  `.pptx` behind for the two to disagree.
- A PDF cannot be checked for fonts, spacing or stray boxes, because it does not record
  them. The app lists the checks it skipped rather than staying quiet about them.

### Choosing your exam board

The mark scheme is selected in **Settings → Exam board and qualification**:

| Board | Qualification | Criteria | Marks |
|---|---|---|---|
| WJEC Eduqas | A-level Design and Technology (NEA) | 5 | 100, published per criterion |
| AQA | GCSE Design and Technology (8552) | 6 | 100: 10 + 10 + 20 + 20 + 20 + 20 |

**One board is loaded at a time.** Sending both mark schemes on every scan would
cost you twice as much and fill the coverage grid with criteria from the wrong
exam. Changing the board clears the AI's previous judgement, because it was made
against the other board's criteria, so the next check re-reads your pages.

The AQA criteria names, their per-criterion marks and their band boundaries come
from the official AQA specification, Version 1.2 (6 June 2022), sections 4.4.4.1
to 4.4.4.6. AQA marks by **best fit within a band** — work that fully meets a band
sits at the top of it, mostly meets it in the middle, just meets it at the
bottom — and the specification states that each band must be viewed
**holistically**. The app's prompt tells the model to place the work in a band on
that basis, rather than inventing a score.

### What it deliberately does NOT do

- It does not write or rewrite your text.
- It does not edit, move, or "improve" your PowerPoint.
- It does not submit anything anywhere.
- It does not do the work for you. That would breach the regulations.

---

## What leaves your computer

**Without an AI key: nothing at all.** The app makes no network requests.

**With an AI key:** the extracted *text* of the slides being reviewed, plus your
mark scheme, is sent to Google's Gemini API. Slide images are **not** sent in this
version. Your `.pptx` file itself is never uploaded.

If that is a concern, do not add a key. The free checks still work fully offline.

---

## What a free API key costs: nothing, but there is a daily allowance

The Gemini free tier needs no card and no payment. What it spends is **requests**,
and Google caps them **per model per day**, resetting at **midnight Pacific**.

Measured on a real free key: **20 requests per model per day**, and one full check
of a 60-page portfolio uses about **9** — roughly **two full checks a day**. Repeat
checks are much cheaper, because pages that have not changed are never re-sent.

So the app does not show a bill. The card at the bottom shows **the percentage of
today's allowance you have used, per model**, and when it resets — because on a
free key a money figure would read £0.00 forever and tell you nothing.

Google does not publish the per-model daily numbers (its rate-limits page points at
AI Studio for the live values), so the app starts from a conservative estimate and
**replaces it with the real figure the first time a check is refused**, when the
error body states the actual limit. An estimated limit is labelled as such, so a
guess is never presented as a fact.

## Cost

**Nothing, on a free key.** The app runs locally: no subscription, no server, nothing
rented. The AI review is billed by Google to your own key, and on the free tier that bill
is always zero — what you spend is part of the daily request allowance described above.

So there is no money anywhere in the app. The card at the bottom shows **the percentage of
today's allowance used, and when it resets**. A running total of £0.00 would tell you
nothing about the thing that actually runs out.

| Action | What it uses |
|---|---|
| Any free structural check | nothing — no request is made |
| Full AI read of 60 pages | about 9 requests |
| "Check now" when nothing has changed | nothing — unchanged pages are never re-sent |
| "Check now" after editing a few pages | the changed pages only, plus one whole-portfolio check |

At roughly **20 requests per model per day**, that is about **two full checks a day** — and
far more on days when you are only editing, because pages you have not touched are never
sent again.

**On a paid key** the same checks are billed in money instead, and the app quotes the cost
before you press the button. Nothing in the interface requires you to pay.

What keeps the allowance lasting:

- **Fingerprint caching.** Every page is fingerprinted, so unchanged pages are never
  re-sent and pressing "Check now" repeatedly uses nothing.
- **The cost is shown before you spend it.** The line above the Check button says how many
  requests the next check will take and how much of the allowance is left.
- **Batching.** Pages are read in groups of 12, which amortises the mark scheme that has to
  be re-sent with every request.
- **Deep thinking off by default.** It is slower and uses more per request, so it is a
  deliberate choice rather than always-on.
- **Swapping keys.** Each key has its own daily allowance, and Settings lets you paste a new
  one in seconds — so running out is a two-click problem, not a wait.

### Making the allowance last

The app already does the two things that matter most, and both are automatic:

- **a cheap model reads the pages, while a stronger one only decides the mark.** Reading
  is the bulk of the work and the easy part; judging is where quality shows, so the
  expensive model is used for the smallest share of it.
- **unchanged pages are never re-sent**, so checking repeatedly while you edit costs a
  fraction of a full check.

**Do not switch the marking model to a Flash-Lite one to save allowance.** It was measured
against three portfolios with known marks and it does not merely read low, it
**compresses** — a 99, a 96 and an 89 scored 81, 80 and 66. It is cheaper and worse in a way
you cannot correct for. Switching **Predict a mark** off in Settings is the honest saving:
you lose the number and keep the full fix list.

---

## Which AI model

Two models are used, and the split is deliberate:

- **Reading the pages** uses a cheap model (`gemini-3.1-flash-lite` by default). Reading is
  the bulk of the requests and the least demanding part.
- **Deciding the mark** uses a stronger model (`gemini-3.5-flash`). This is the setting that
  matters: measured against three portfolios with known marks, the cheap model scored 81,
  80 and 66 where the truth was 99, 96 and 89 — compressed far worse than the strong model's
  2-to-6-marks-low. Change it under **Settings → Stronger model for marking**.

The daily allowance is counted **per model**, so the two do not compete for the same
requests — and a key that has run out on one still has room on the other.

If Google renames or retires models, the app **asks the API which models your key can
actually use** and picks a sensible equivalent automatically. It ignores models Google still
lists but which refuse new keys with a 404.

To pin a specific model, set `"model"` in `<your data folder>/config.json`.

---

## Settings

Most settings are changed in the app itself, under **Settings**:

- **Your portfolio** — one picker that takes a `.pptx`, a PDF export, or both. They are not
  interchangeable: the PDF is what the AI looks at (every page, including drawn work), and the
  `.pptx` is what the layout checks can measure (uncaptioned images, text off the page, speaker
  notes). The section itself says which is which.
- **Exam board and qualification** — which mark scheme to judge against.
- **Model** — which AI model reads your work, with the requests per check beside each.
- **Predict a mark** — on by default. Off makes a check about a third cheaper and
  still gives the full fix list, because the mark is the only part that needs
  repeated runs so one odd answer cannot decide it.
- **Stronger model for marking** — optional. Used only for the marking runs.
- **Colour theme** — match your computer, or force Light or Dark.
- **Your Google API key** — paste a new one, test it, or remove it. Testing a key uses no
  allowance, so it is safe when you have run out.
- **Use it on your phone** — its own section at the foot of the page, with the current address
  to type on the phone including the access key, and a "Check again" button for when the router
  has reassigned it. It was previously buried in the advanced settings, which is the wrong place
  for it twice over: it is not an option to tune, and it was the one part of the page a phone
  user could not find on the phone they were holding.
  code. Shown only on the computer running the app, never sent to another device.
- **Deck folder path** — only needed if you would rather keep your file where it
  is than let the app copy it in.

Under the hood they live in `<your data folder>/config.json`:

```json
{
  "deckDir": "/path/to/your/portfolio/folder",
  "deckFile": null,
  "pdfFile": null,
  "specId": "eduqas-alevel",
  "model": "gemini-3.8-flash",
  "visionModel": "",
  "markEnabled": true,
  "apiKey": "",
  "watch": true
}
```

- **`deckDir`** — the folder to look in. If it contains several `.pptx` files, the
  most recently edited one is reviewed. Change this to point at your real portfolio.
- **`deckFile`** — set a filename to pin one specific file.
- **`pdfFile`** — the PDF used for the visual review. Set by uploading one in Settings.
- **`visionModel`** — a stronger model used only for the marking runs; blank means
  the same model as `model`.
- **`markEnabled`** — `false` skips the repeated marking runs, at about a third less
  per check, and simply omits the predicted mark.
- **`watch`** — when true, the app notices when you save changes in PowerPoint and
  can re-scan automatically. It waits a few seconds after a save so it does not
  scan a half-written file.

You can also change the folder and model from the Settings section in the app.
The app **never** displays or transmits your API key.

---

## Your mark scheme (this is already set up)

Rohan's real mark scheme is **already loaded**, extracted from the official
WJEC Eduqas A Level Design and Technology specification (Version 2, January 2019,
Appendix A). The app is judging his work against these five criteria, worth
100 marks in total:

| | Criterion | AO | Marks |
|---|---|---|---|
| (a) | Identifying and investigating design possibilities | AO1 | 15 |
| (b) | Developing a design brief and specification | AO1 | 15 |
| (c) | Generating and developing design ideas | AO2 | **25** |
| (d) | Manufacturing a prototype | AO2 | **25** |
| (e) | Analysing and evaluating design decisions and prototypes | AO3 | 20 |

AO1 = 30%, AO2 = 50%, AO3 = 20%, which matches the specification's own weighting
table. Note that **(c) and (d) together are half the marks**, so the app pushes
hardest on those.

The criteria live in `spec/eduqas-dta-nea-criteria.json`, which also carries the
**Band 5 (top band) descriptors** for each criterion. That is what teaches the AI
what the examiner actually rewards, rather than generic advice.

`spec/_YOUR-MARK-SCHEME.md` is a human-readable version of the same thing — the
official band thresholds and what each criterion is asking for. It starts with an
underscore so it is **not** sent to the AI; it exists for you and Rohan to read.

### How spec files are chosen

- Files starting with `.` or `_` are **ignored** — that is how you park a file
  without it being loaded.
- Documentation-style names (`README`, `GUIDE`, `YOUR-`, `example`, …) are ignored.
- A structured `.json` is **authoritative**. If one is present, headings from
  `.md`/`.txt` files are not also turned into criteria, so you never pay to send
  the same mark scheme twice.
- With no `.json`, criteria are inferred from Markdown headings — which is enough
  for a plain-text spec.

### Swapping in a different spec

Delete or rename the existing criteria file, then drop in your own `.json`:

```json
{
  "criteria": [
    { "id": "c1", "name": "Client and problem definition", "weight": 10,
      "ao": "AO1",
      "description": "Identification of a client and a justified design problem.",
      "evidence": "What the top band descriptor says." }
  ]
}
```

**PDFs are not read directly.** The app reports a skipped PDF rather than silently
ignoring it. The Eduqas spec was converted from PDF using `pdftotext`, and the
criteria were extracted from Appendix A.

---

## Troubleshooting

**"No .pptx file was found"**
Your portfolio is not in `deckDir`. Put the file there, or change `deckDir` in
Settings. The path can be found in Finder with right-click → "Copy as Pathname".

**Browser says it cannot connect**
Give it a couple of seconds. If it persists, the server did not start: read the
message in the terminal window.

**"Port 4173 is already in use"**
NEA Coach is probably already running. Open <http://127.0.0.1:4173>.

**The AI says the key was rejected**
The key was mistyped or has a trailing space. Copy it again from Google AI Studio.

**The AI says it is rate-limited**
Wait a minute and press Check now again. Nothing is lost: pages already read are
cached, so the retry costs almost nothing.

**Findings look wrong or too harsh**
The thresholds are in `lib/checks.js` at the top, with comments. They were tuned
for A3 pages and are easy to adjust.

**Everything is slow**
The free checks take about a second for 200 pages. Only the AI pass is slow,
because it is waiting on the network, not on your computer.

---

## For whoever maintains this

```
NEA-Coach/
├── Start NEA Coach.command   Double-click launcher (macOS); fetches Node if needed
├── Start NEA Coach.bat       Double-click launcher (Windows); needs Node installed once
├── tools/calibrate.js        Measures the predicted mark against known marks
├── server.js                 Local HTTP server, scan orchestration, report
├── lib/
│   ├── zipread.js            Minimal read-only ZIP reader (no dependencies)
│   ├── pptx.js               OOXML parser: text, fonts, images, autofit, notes
│   ├── imagemeta.js          Image dimensions from file headers
│   ├── checks.js             The free deterministic rules
│   └── ai.js                 Gemini client, batching, fingerprint cache, costs
├── public/index.html         Entire frontend, self-contained, no CDN
├── sample/                   A 24-page test deck with deliberate flaws
├── spec/                     Drop your mark scheme here
└── runtime/                  Downloaded Node binary (created at runtime)
```

**Where your data lives**

Nothing you care about is stored inside the app folder. Your settings, your API key,
the saved report, the scan history and the copy of your portfolio all live in a
per-user data folder, so unzipping a new version beside the old one changes nothing:

| Platform | Where your data is |
| --- | --- |
| **macOS** | `~/Library/Application Support/NEA Coach` |
| **Windows** | `%APPDATA%\\NEA Coach` |

The folder is created on first run. If you are upgrading from v1.0.22 or earlier, the
app brings your existing settings and portfolio across the first time it starts, and
says so in its window. Set `NEA_COACH_DATA_DIR` to put the folder somewhere else,
which is also how you would keep the app and its data together on a USB stick.

To start again from nothing, delete that folder. To find your saved reports, they are
the `*.json` files in it.

**Putting it on a phone, or on the desktop**

- **Phone.** Scroll to **Use it on your phone** at the bottom of the page, and open that address
  on the phone — the app works it out from the Mac's own network interfaces, so it is always
  current. Then
  use your browser's *Add to Home Screen*. The name and icon come from the app itself, so it
  looks like an app rather than a bookmark. On iOS use **Safari** — Chrome and Firefox on iOS
  cannot add to the home screen.
- **Mac.** The download contains `NEA Coach.app`. Drag it to the Desktop or to Applications.
  It carries its own copy of everything, so it keeps working after being moved.
- **Windows.** `Start NEA Coach.bat` puts a shortcut with the icon on your Desktop the first
  time it runs. Deleting the shortcut is respected — it is not recreated on every launch.

The icons come from one drawing in `tools/make-icons.js`, so the browser tab, the phone
icon, the Mac app and the Windows shortcut cannot drift apart into four near-misses. That
script only needs running if the design changes; its output is committed.

**Design constraints, deliberately chosen:**

- **Zero npm dependencies.** Nothing to install, nothing to break, no supply chain.
  Everything uses Node built-ins. This is why it works identically on Intel and
  Apple Silicon.
- **No build step.** The frontend is one self-contained HTML file.
- **Two independent detection layers.** `checks.js` is deterministic and offline;
  `ai.js` does judgement. Failures in one never take down the other.
- **Never trust a heuristic without a render.** The `text-overflow` rule is the
  worked example: an early version flagged a caption that a real LibreOffice
  render proved fits, which is why `autofit` is now parsed and three distinct
  cases are handled separately.

**Known environment issue:** the bundled LibreOffice Kit's direct PPTX→PNG path
can emit fully transparent pages. Converting PPTX→PDF first and rasterising that
works correctly, which is the route the sample renders used.

**API endpoints:** `/api/state`, `/api/progress`, `/api/scan` (POST),
`/api/config` (GET/POST), `/api/models`, `/api/health`, `/api/log`.

---

## On your iPhone

**This works now — it is not a separate app.** The same interface opens in Safari
on the phone, laid out for a small screen. Nothing to install.

The server listens on all interfaces (`"host": "0.0.0.0"` in `config.json`),
so the phone can reach it. When you start the app, the launcher window prints
the exact address to type:

```
   On this Mac:  http://127.0.0.1:4173

   On your phone (same Wi-Fi, or Tailscale from anywhere):
     http://192.168.0.87:4173/?key=abcd-efgh-jkmn   same Wi-Fi (en0)
   The key in that address is the password. Do not share it.
```

Type that second address, **including the `?key=` part**, into Safari on the
phone. Add it to the Home Screen (Share → Add to Home Screen) and it opens
full-screen like an app, with the key remembered.

### The access code

Only this Mac gets in without a code. Every other device needs the key in the
address, which is stored as `accessCode` in `<your data folder>/config.json` and stays the
same across restarts. Without it, anyone on the same Wi-Fi could read the
report, replace the API key, point the app at any folder on the Mac and run
scans that cost money. Web pages open on the Mac are blocked from changing
settings or starting scans too.

If the code ever leaks, delete the `accessCode` line from `<your data folder>/config.json`
and restart: a new one is made. To make the app Mac-only, set `"host"` to
`"127.0.0.1"`.

### What the phone can do

Everything: read the Fix list, expand the Worth adding tab, view the Coverage
grid, browse the slides table, mark items done, and **press Check now to run a
new scan**. The scanning happens on the Mac; the phone is just the screen. Keep
the Mac awake while scanning from the phone.

### The API key works from the phone

The AI features work fully from the phone. The key itself **stays on the Mac** and
is never sent to any client — the app only reports whether a key is set
(`apiKeySet: true/false`). Verified: no API endpoint returns the key, and no log
line contains it.

### On a small screen the layout adapts

A 103-item fix list is about 50,000 pixels of scrolling on a phone, which is
unusable. On screens under 700px wide:

- The Fix list shows the **top 5 items**, with a "Show 10 more" button and a
  "Show all" option. Nothing is hidden permanently.
- The **filter and search controls fold away** behind a "Filter and search"
  summary, since they are not what you want to see first.
- Cards, headings and tabs are tightened so the first screen is readable.

All three tabs, the cost meter, the slides table and Settings remain available.

### From anywhere, not just at home

Home Wi-Fi only works while you are in the house. For access anywhere, install
**Tailscale** on the Mac and on the iPhone (free, and nothing is published to the
public internet). Both devices join the same private network, and the `100.x`
address the launcher prints will then work from anywhere.

### If the phone cannot connect

**"Safari can't open the page" almost always means the address changed.** Your Mac's
address on the Wi-Fi is handed out by the router and can change; when it does, an
address that worked yesterday simply will not resolve. This is the single most
common reason the phone stops working, and it is not a fault in the app.

**The current address is in the app**, under **Settings → Use it on your phone**.
Copy it from there rather than reusing an old bookmark. For the same reason, the
launcher also prints it every time it starts, so the terminal window is the other
place to look.

The `?key=` part is the password. Without it the app refuses any device that is
not the computer running it, which is deliberate — it stops anything else on the network reading
your report or spending your API credit.

If the address is right and it still will not connect:

1. Check the phone is on the **same Wi-Fi** as the Mac, not mobile data.
2. Check the Mac has not gone to sleep.
3. Use the exact address, including `:4173`.
4. Some home routers have "AP isolation" or a guest network that blocks devices
   from talking to each other even on the same Wi-Fi.

**For a permanent fix, install Tailscale** (free) on both the Mac and the phone.
It gives a `100.x` address that never changes and works from anywhere, not just at
home, which removes this problem entirely. The app shows it when it is available.

---

Version 1.0.0
