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

- **The page is eleven sections behind a tab strip along the top.** No more scrolling to the
  bottom for Settings, or hunting for the coverage grid:
  **Overview · Fix list · Things to add · Predicted mark · Coverage grid · Since your last check ·
  Slides · Statistics · Past checks · Information · Settings**
- The order is the one I would use it in: **the work first** (overview, then the two lists you act
  on), then the mark, then the reference material, with **Settings last** because you visit it
  once. Say the word if you would rather something sat elsewhere — it is a one-line change.
- On a phone the strip is **one row you swipe sideways**. Eleven tabs do not fit across a narrow
  screen, and wrapping them onto rows would push the content off the bottom of it.
- **"Use it on your phone" is now at the top of Information**, and **the allowance meter is in
  Information too**, as asked.
- The open tab is the accent colour, and the strip scrolls it into view — with eleven tabs the one
  you are on is usually off the right-hand edge.

**Two bugs that only looking at the page found**

- **The footer said "Version 1.0.0".** It was a hardcoded constant that no release ever updated, so
  it had been wrong for fifty versions. It now shows the version the running app reports, which is
  the same value the update notice uses, so the two cannot drift.
- **The footer jumped to the top of the page** once the sections became panels, because the page is
  now a flex column and the footer's default position was first.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
