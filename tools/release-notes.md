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

- **The "Worth adding" suggestions have colour now.** It was a white card of grey text with three
  small pills on it, which is why it read as the drab one:
  - each suggestion carries a **pale accent wash**, so the section is visibly its own thing
  - the **number badge is solid accent** instead of the fix list's grey
  - **"Why it is worth it" and "What to do"** are accented — and it turned out those two labels
    had no styling at all. `.finding-block` is a class the app emitted and the stylesheet had
    never heard of, so the two lines that give a suggestion its structure were plain bold text.
- Suggestions deliberately do **not** use the Fix list's severity colours: nothing there is a
  mistake. Green still means "quick and worthwhile", grey still means "small gain", so the
  colours that carry meaning are unchanged.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
