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

Four more things move in the background, on top of the drape bands:

- **Dust motes** drifting slowly upward — tiny specks, like dust in sunlight.
- **Colour clouds** — two big soft ones, well behind everything and much slower, so the background
  has some depth to it.
- **Two soft glows** that slowly brighten and fade.
- **Once in a while, a faint streak of light** crossing the page. It spends most of its cycle
  invisible; the crossing takes about five seconds out of every thirty-four.

**Still built so it cannot make the page stutter.** Every one of them animates *transform* or
*opacity* only, which the graphics card handles without touching the main thread or repainting
anything. There are seven moving layers in total, and a test fails if that rule is broken or if the
count grows without someone deciding to.

The dust is worth a note: it is **one** layer carrying a tiled field of specks, scrolled by exactly
its own height so the loop is seamless. Twenty individually animated dots would have been twenty
layers to composite for the same picture, and would have cost more than everything else here
together.

It is all blue and teal — the violet and pink are gone, since the background is meant to stay soft
blue.

**Next:** the tabs. Overview, Fix list, Things to add, Predicted mark, Coverage grid, Since your
last check, Slides, Statistics, Past checks, Information, Settings — across the top, in one
scrollable row on a phone, with "use it on your phone" moved into Information. That is a large
restructure of the page, so it is getting its own release rather than being rushed alongside this.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
