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

Five faults found by a full read-through of the code, all of the same shape: something that made
a feature work was missing, so the feature existed, logged that it had done its job, and did
nothing.

- **Cancelling a check no longer throws away the pages you had already paid for.** The partial
  cache was overwriting the good one, so the next check re-sent everything after the point you
  cancelled at. The previous entries are carried through now.
- **The daily allowance can finally show a percentage.** The app learns your real daily limit
  when Google tells it — but that figure was being dropped as it was read, and then deleted by
  the next request, so the meter could only ever show a count. It is kept.
- **"I could not read the document" is no longer treated as a successful review.** If the model
  reported that it could not see your pages, the app still recorded that they had been looked at
  and reported a mark from them. It now says so instead.
- **Two things were nearly invisible in dark mode**: the little counts on the severity buttons
  (about 1.1:1 against the background, where 4.5:1 is the minimum) and the "what is already
  working well" text (about 1.6:1). Both follow the theme now.

**In v1.0.40**

- Colour, much less work while idle, and Cancel actually stopping the visual review.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
