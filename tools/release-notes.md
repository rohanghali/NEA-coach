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

- **The page has colour in it now.** It was white boxes on a white page; the headings carry the
  accent colour and the background is a soft wash that the cards let through. The colour coding
  of the findings is untouched, and nothing is louder than it needs to be.
- **Cancelling a check now actually stops it.** The Cancel button could not stop the visual
  review — the first and far and away the most expensive thing a check does, up to three
  whole-portfolio reads plus a backup model — because the flag was never passed to it. Pressing
  Cancel did nothing until every one of those had already run and been paid for. It stops before
  the first one now.
- **Much less work while sitting still.** The "Check now" button used to pulse for ever, and the
  pulse was a shadow — which cannot be animated on the graphics card, so it repainted the button
  sixty times a second, permanently, for decoration. It invites three times and then rests. The
  background's gentle lean towards your pointer no longer animates either; it was repainting the
  whole screen continuously while the mouse moved.

**In v1.0.39**

- **Fixed the page stopping halfway through drawing itself** — an SVG element refuses a
  `className` assignment, and that one line was taking the whole render down with it.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
