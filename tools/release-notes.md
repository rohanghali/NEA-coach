Reviews a D&T NEA portfolio and tells you what to fix, in order of what is costing you marks.

**To use it:**

**Windows** — download `NEA-Coach-windows.zip`, unzip it, then double-click
`Start NEA Coach.bat` inside the folder it creates.

**Mac** — download `NEA-Coach-mac.zip`, unzip it, then double-click
`Start NEA Coach.command` inside the folder it creates.
(The first run downloads Node.js automatically; after that it opens instantly.)

Each download unzips into a folder named after its version, so unzipping a newer one
never lands beside an older one or overwrites it.

Nothing needs installing by hand, and no bank card is needed.

Get a free Google API key — the app explains it, and `SETUP-AI.md` is in the folder.

**New in this version**

- **Fixed the big one: a portfolio can no longer be marked low just because the marking model
  was busy.** Google occasionally returns "this model is currently experiencing high demand".
  When that happened, the app gave up on looking at your pages entirely and reported a
  **text-only** figure as though it were a mark — 74 out of 100 for a portfolio that really
  scored **99**. It now moves to a different model and looks at the pages anyway.
- **It can only move to a model it has been checked against.** The app has compared itself with
  three real marked portfolios per model, and it will only fall back to one of those — because a
  model with no measured error cannot be corrected, and an uncorrected number is exactly what
  went wrong here.
- **If a review still cannot happen, it now says so next to the number.** It names the cause
  ("the marking model was busy", "you have used the daily free allowance"), says plainly that the
  figure comes from your text alone and is not a judgement of the work, and tells you to press
  Check now again. A plausible-looking wrong number was worse than an honest gap.
- **The correction now follows the model that actually did the marking.** Previously it used the
  model you had *chosen*; if a fallback had been used, the wrong line — or none at all — would
  have been applied.

**In v1.0.29**

- **An icon**, in the browser tab, on a phone home screen, and on the desktop, plus a real
  `NEA Coach.app` for macOS and a Desktop shortcut on Windows.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
