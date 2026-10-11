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

- **The allowance now counts each model separately, which is what was wrong with it.** Your free
  allowance is per model, but every request the app made was recorded against the *text* model —
  including the page-reading requests, which use the marking model. So the text model looked
  busier than it was, the other model's row read zero, and neither matched what Google actually
  shows you. The number of requests was right; which model they belonged to was not.
- **"Worth adding" is no longer the grey one.** Unselected tabs were pale grey on a white card,
  which read as disabled rather than as *"there is something in here"*. They are normal ink now,
  the open one is washed with the accent, and the suggestions panel has the accent as its own
  colour — suggestions are not mistakes, so they should not wear the Fix list's severity colours.
- The focus ring around whatever you are tabbing to was a pale blue in four places, which is weak
  on a white card. It is the accent colour now.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
