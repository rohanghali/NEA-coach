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

- **The tab scroller is a much lighter, transparent blue.** Solid, it read as a filled progress bar
  — something telling you a number — when it is a handle you are meant to grab. See-through, it sits
  *on* the rail rather than replacing it, and the track stays visible behind it. It firms up a little
  under the pointer so it is obvious you have hold of it.
- **The Overview note now says whether it will MARK your work, and whether DEEP THINKING is on:**

  > The app would run 13 of the 18 free checks on your 60 pages, read every page, and **mark** the
  > work against **AQA GCSE Design and Technology (8552)**. **Deep thinking is on** — it looks harder
  > at the work, takes longer, and uses more of the allowance.

  With no key it says **"It would NOT mark the work"** in amber instead, because that is the
  difference between a mark and no mark.

- **Fixed: the board was showing as a file name.** The note said *"mark the work against
  aqa-gcse-dt-criteria.json"* — and so did the predicted mark, the Settings summary and the coverage
  grid's empty state, in four places. The app was always written to prefer a proper label; the server
  simply never sent one. It does now, so all four read **"AQA GCSE Design and Technology (8552)"**.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
