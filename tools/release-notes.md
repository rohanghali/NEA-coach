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

A full audit of the whole app turned up thirteen real bugs. These are the ones that affected what you
were told, fixed together.

**Findings that were wrong**

- **Two spellings that are both correct English are no longer reported.** The rule offered to "correct"
  **licence/license** (noun vs verb) and **metre/meter** (unit vs measuring instrument). It also matched
  any longer word *starting* with the American form, so writing about a **colorimeter** triggered a
  "you mixed in American spelling" finding for a word you never wrote.
- **Misspellings are no longer deleted.** Every spelling finding was merged into one row, so only the
  first was ever shown — and that row listed the pages of the *other* typos. On your portfolio this is
  why the report now says **31 things to look at instead of 30**: findings that were being silently
  dropped are now their own rows, each with its own pages.
- **The allowance note no longer claims checks that cannot run.** Two rules are switched off, and they
  were counted anyway — the app said "all 18 free checks" and told you your image resolution had been
  checked. It says 16 now, and 11 for a PDF.

**Pages that were never actually reviewed**

- **A batch the AI could not read was recorded as reviewed and cached as clean.** If the model said it
  could not assess a batch, the empty result was stored against those pages and served from cache on
  every later check — up to 12 pages getting no advice at all while the report counted them as read.
  It now fails the batch and retries it next time.
- **A replaced picture is a changed page.** The fingerprint used the image's *file name* rather than its
  contents, so swapping or recompressing a photo left the page "unchanged" and the old review was
  reused. It now uses the archive's own record of the image, which costs nothing on a large deck.

**Settings that were not being saved**

- **Deep thinking was never saved.** The switch worked for that one check, but the setting never
  reached the server — so the front page said "Deep thinking is off" while your switch said on, the
  choice was lost on reload, and automatic re-checks ran without it.
- **Cost estimates used the previous portfolio's page count**, and one of them assumed 60 pages when
  it knew nothing. After swapping files the app priced the wrong document.

**Smaller things**

- The first-run wizard asked whether a portfolio was added by reading a field the server never sent, so
  a **.pptx-only** setup was told it had added nothing.
- The **welcome-back note could never appear** — it was hidden with an attribute that "show" did not clear.
- All eleven tabs pointed at panel names that did not exist (screen-reader links).
- The coverage grid told you to drop in a **PDF** of the criteria; the loader skips every PDF. It asks
  for **.txt or .md** now, and says why any file it found was not used.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
