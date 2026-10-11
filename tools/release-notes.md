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

- **The background actually moves now — you were right that it did not.** It was moving, technically:
  about **a pixel a second**, which is indistinguishable from still. The bands crossed roughly 3% of
  their width over 54 seconds.
- **Now: about four times faster and three times further.** A pass takes 19, 27 and 37 seconds, each
  band travels about a tenth of its own width — roughly 150px on a laptop — and leans up to three
  degrees, which swings its ends by about 100px. The three move in **different directions**, so they
  read as folds crossing rather than one sheet sliding.
- **This costs nothing.** A composited transform is the same work whether it moves a little or a
  lot, and fast or slow: the layer is drawn once and the graphics card re-places it. The rule that
  keeps it cheap — transform only, never a blur or a shadow — is unchanged, and tested.
- If it still looks still to you, check **System Settings → Accessibility → Display → Reduce motion**.
  That setting switches the whole background off by design, and it is meant to.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
