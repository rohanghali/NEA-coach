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

Nothing you see in the report has changed. This is about **starting** the app.

- **The app opens your browser itself, at the port it actually got.** Both launchers used to do
  it, by waiting for a port they had assumed was 4173. When that port was busy the app moved to
  the next free one and remembered it — so the launcher waited on the wrong port, printed *"did
  not start within 30 seconds"* while the app was running perfectly well next door, and never
  opened anything. The app knows its own port; it opens the browser the moment it is listening.
- **An older copy is now called out clearly.** If a different version is already running, the app
  says which one it found, which one this folder is, and how to switch — instead of silently
  opening the old one. On Windows this check never worked at all: it was written in a way that
  cmd silently breaks on any path containing a space, which `C:\Program Files\nodejs\` always
  does.
- **The launchers are much smaller.** They start the app and stop. The port waiting, the browser
  opening and the version comparison all live in the app now — one implementation instead of two,
  and it cannot disagree with itself.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
