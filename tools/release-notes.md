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

- **Fixed: clicking Settings scrolled the tab strip to the end.** The strip tried to bring the open
  tab into view, but it measured each tab's position from the *page* rather than from the strip — and
  Overview and Settings are pinned **outside** the strip, so Settings' position out at the far right
  dragged the whole middle row to its end every time. Only a tab that lives in the strip moves it now,
  and the distance is measured from the strip itself.
- **Fixed: the caret on the Settings heading.** Settings is its own section now, so there was nothing
  for it to fold — but a small ▾ was still drawn on it and did nothing when clicked. That is the
  "drop down" that was not working. It is gone. The inner groups *within* Settings (your portfolio,
  the board, the AI key, and the tuning options) still fold, because those genuinely are sections.
- **Statistics now sits after Past checks**, as asked.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
