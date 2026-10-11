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

- **"Use it on your phone" is now its own section at the bottom of the page**, instead of being
  buried in the advanced settings. It was in the wrong place twice over: it is not an option to
  tune, and it was the one part of the page a phone user could not find on the phone they were
  holding. It now says to use *Add to Home Screen* once it opens.
- **A "Check again" button on it.** The address changes whenever your router reassigns it, so
  "come back here and copy the new one" needed a way to actually come back. It also fills itself
  in as soon as the page loads, rather than waiting to be asked.

**In v1.0.30**

- **A busy marking model can no longer turn a 99/100 portfolio into 74.** When Google returned
  "this model is experiencing high demand", the app gave up on reading your pages and reported a
  text-only figure as a mark. It now moves to a different model and looks at the pages anyway —
  and only ever to one it has been calibrated against, so the result can still be corrected.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
