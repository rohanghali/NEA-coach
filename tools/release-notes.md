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

- **Fixed a button that had quietly gone missing, and 528 MB of disk space with it.** Uploading
  a PDF clears the PowerPoint reference, and the app then stopped believing it was holding one —
  so the **Remove PowerPoint** button disappeared while its own copy of a 528 MB `.pptx` sat in
  its folder, mentioned nowhere in the interface. That copy is now visible, and the button says
  what it will delete and how big it is.
- **"Remove PowerPoint" now actually removes it.** It used to clear the reference and leave the
  file on disk, which is how the copy survived. It deletes the app's own copy now — and only
  ever that. Your original file is never touched.
- **A leftover copy is shown for what it is**: *"Rohan Ghali.pptx (528 MB) stored, but you are
  being marked on the PDF — the button below removes it"*, flagged rather than ticked as though
  it were in use.
- **The list can no longer call a PDF a PowerPoint.** It once showed the PDF's name in the
  PowerPoint row, because a PDF-only setup leaves no PowerPoint name to use and the row fell
  back to whatever had been marked.

**In v1.0.32**

- A build with no student names in the shipped source. Nothing behavioural.

**In v1.0.31**

- **"Use it on your phone" is its own section at the bottom of the page**, with a "Check again"
  button for when your router reassigns the address.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
