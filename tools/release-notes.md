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

- **Fixed the app appearing not to start at all.** If the port it wants is already taken, the app
  used to print one line and close — which looks exactly like the app being broken. That happens
  in the two most ordinary situations there are: opening it twice, and installing an update while
  the old copy is still running, because every version asks for the same saved port.
- **Now it works out which case it is.** If NEA Coach is already running, it says so and opens
  that copy — your work is in the window that is already open, so there is no reason to start a
  second one. If something *else* is on the port, the app moves to the next free one and
  remembers it.
- Nothing about the report, the marking or the checks has changed.

**In v1.0.36**

- **A Cancel check button**, which stops at a safe boundary and keeps every page already read, so
  the next Check now carries on rather than starting over.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
