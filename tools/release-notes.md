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

**Ask AI actually works now.** It did not: every question came back with an internal error, because a
setting the code needed was not loaded where that request runs. Two more of the same kind were hiding
behind it — one broke the general questions, one broke the per-criterion ones — so **every kind of
question** is checked end to end now, not just the ones I happened to try.

- **It is clear you can type your own question.** The section says so, and the suggestions are
  introduced as a starting point rather than the whole list.
- **Copying is instant.** The file is built quietly in the background when you open the section, and
  the app keeps it once built, so pressing Copy is immediate instead of a several-second wait. It is
  rebuilt automatically if you change your portfolio.
- **Copying and opening ChatGPT are separate buttons.** One press used to do both, which pulled you
  away from the page while it was still building the file, with no way to tell whether it had worked.
  There is now **Copy everything for ChatGPT**, **Open ChatGPT**, and **Save as a file**.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
