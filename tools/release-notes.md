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

- **The app now tells you if its page and its brain are different versions.** That happens in one
  specific way: unzipping a newer copy on top of one that is still running. The running copy
  keeps its old code in memory while every page it hands out is read from the new files on disk,
  so half of it is new and half is old — some things work and some do not, and nothing on screen
  says why. It is now stated plainly, with what to do about it.
- **Starting the app when its port is busy no longer looks like a crash.** If NEA Coach is
  already running it says so and opens that copy. If something else has the port, it quietly
  moves to the next free one and remembers it. Before this, both cases printed one line and
  closed the window.

**If you are updating: quit NEA Coach completely first** — close its window, or press Control-C
in it — and then open the new copy. Updating on top of a running copy is the one way to end up
with two versions running at once, which is what this release is about.

**In v1.0.37**

- The port fix above.

**In v1.0.36**

- **A Cancel check button**, which stops at a safe boundary and keeps every page already read.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
