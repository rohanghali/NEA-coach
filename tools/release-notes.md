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

- **Fixed the page stopping halfway through drawing itself.** One line set a class on the score
  ring, which is an SVG element — and an SVG element refuses that, so it threw and took everything
  below it down with it. The parts drawn before it appeared, the rest stayed blank, and a
  "something went wrong" notice stayed on screen for the rest of the session. **This is almost
  certainly what you saw as the app not working.**
- **That notice now clears itself** once a page draws successfully, instead of staying for good.
- **The update notice from v1.0.38 no longer gives a false alarm.** It compared the page against
  the version that wrote the *report* on disk, which after an upgrade is the previous version
  until the next check is run — so it warned about a mismatch that was not there. It compares
  against the running app now, and the report's own version is kept separately.
- Nothing about the marking, the checks or the report has changed.

**In v1.0.38**

- The port fix: a busy port no longer closes the window.

**In v1.0.36**

- **A Cancel check button**, which stops at a safe boundary and keeps every page already read.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
