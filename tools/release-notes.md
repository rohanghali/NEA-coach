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

- **The background moves now** — three soft drape bands, like folded cloth, drifting very slowly
  over each other behind the page. Still pale blue and still quiet: it is meant to be noticed over
  a minute, not in a glance.
- **It is built so it cannot make the page stutter.** Every one of the three layers animates
  *transform* only, which is the one kind of animation a browser does on the graphics card without
  touching the main thread or repainting anything. No blur, no shadows, no sliding gradients — all
  of those force a full-screen repaint sixty times a second, which is what made an earlier version
  of this page feel heavy. Three layers, and a test now fails if a fourth is added or if any
  keyframe animates anything but transform.
- **The old drifting colour wash no longer animates.** It is a flat, static tint now, because there
  is no sense paying for two sets of motion at once — and its 52-second drift was too slow to see
  anyway.
- **It stops while the tab is in the background.** A hidden tab has no reason to keep a graphics
  layer awake.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
