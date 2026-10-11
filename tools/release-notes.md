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

- **The app now looks the way it works.** A slow tint drifts behind the page, cards are
  translucent with a lit top edge, and the header is a glass bar that stays with you —
  so Check now is reachable from anywhere in a long report instead of only from the
  very top.
- **Buttons answer back.** Every button springs under your finger and a ripple spreads
  from exactly where you pressed. That matters more than it sounds: a check is a
  network call, and a button that does nothing visible for two seconds feels broken.
- **The numbers roll up.** The predicted mark and the count on the ring climb into
  place instead of appearing.
- **Ticking off a finding pays you back**, with a small burst of colour where the tick
  was. A finished check gets a bigger one.
- The ring draws itself in from empty, the progress bar sweeps, and the panels rise
  into place as the page loads.
- All of the above is switched off if you have asked your system for reduced motion.

**Carried over from the last few versions**

- Settings save themselves. Change the model, the marking model or Predict a mark and
  it is applied immediately, with a "Saved" confirmation — no Save button to forget.
  The folder path keeps its own button, because that is the one that can be wrong.
- Choosing your exam board during the setup walkthrough updates the board dropdown.
- The predicted mark no longer inflates itself above what the AI actually read. On one
  portfolio it showed 98 against a real mark of 90, when the AI's own reading of 81 had
  been closer. The adjustment is applied only when the app's measured error is tight
  enough for it to be worth applying.
- Marking uses the wording the exam board itself uses for each criterion, including what
  AQA says is typically missing at each level, taken from its candidate record form.
- The criteria table states its own total, so it can no longer appear to contradict the
  headline above it.
- Downloads unzip into a versioned folder, and the launcher tells you if an older copy
  is still running behind the one you just opened.
- Fewer false alarms: a genuinely good portfolio is no longer told its cover page is too
  empty, and near-identical font sizes no longer count as different headings.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
