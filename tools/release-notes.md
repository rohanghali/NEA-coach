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

- **Fixed: the Fix list's filters were showing at the top of "Things to add" and "Coverage grid"** —
  the severity chips (Critical 2, High 5, Medium 20…), the page box, the search box and "Showing 30
  of 30 findings". The filter card had no id of its own, so the app was hiding its *heading* and
  leaving the card on screen. It has an id now, and goes away with the Fix list.
- **"Things to add" is ranked, with no severities on it.** The list is ordered by how much each idea
  would help for the time it takes, and the numbers down the side are that ranking. What has gone is
  the **graded badge** that used to sit on every suggestion — "Would help a lot" or "Small gain", in
  the same greens and greys the Fix list uses for severity. The data was never severity (the prompt
  is explicit that it is how much fixing something would *gain*), but a colour-coded rating at the
  end of every row reads as one whatever it is called. Suggestions are not faults. The time estimate
  stays, because that is effort and not a grade.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
