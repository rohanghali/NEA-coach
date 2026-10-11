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

- **Removed a warning that was not true.** If you supplied only a PDF, the app used to
  say that "the free structural checks have nothing to read and the pages cannot be
  counted". It does not: a PDF is read into a real document with its page count and the
  text of every page, so those checks run and the pages are counted. What a PDF genuinely
  cannot supply — image resolution, shape geometry, speaker notes — is already named
  check by check in the coverage panel, which is the right place for it.

**In the previous version (v1.0.25)**

- **A "Top" button.** Once you are a few screens down a report a small floating button
  appears in the bottom corner and takes you back to the top. It is not there until it is
  useful, and it cannot be reached by keyboard while it is invisible.

**In v1.0.24**

- **The smallest font the app will complain about is now 10pt, down from 12pt.** On one
  real portfolio awarded 90/100, the *only* two mechanical complaints the app made were
  two pages set at 10.6pt. Both are gone.

**In v1.0.23**

- **Nothing is reviewed until you have chosen a portfolio.** Pressing Check now with
  nothing added used to secretly review whatever PowerPoint it found lying around.
- **Your settings, your key and your portfolio now live outside the app folder**, so
  unzipping a newer version keeps all of them.
- **It looks the way it works.** Translucent cards, a glass header bar, buttons that
  spring and ripple where you press, and a small burst of colour when you tick a finding
  off. All of it is off if you have asked your system for reduced motion.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
