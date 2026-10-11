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

**The home-screen icon, properly this time.** My last fix addressed the wrong thing — it corrected a
URL that was returning "not found", which was a real gap but not the reason yours did not change.

- **The icon had transparent, rounded corners.** iOS puts the icon inside its own rounded shape and
  treats transparency as BLACK, so a picture that is already rounded and transparent gets rounded a
  second time and shows dark corners. It now sends a **flat, square, fully opaque** image, the way
  Apple asks for, and lets the phone do the shaping. The same applies to the Android icons.
- **A phone caches a home-screen icon against its address and keeps the old one even after you
  re-add it.** Fixing the picture alone reaches nobody who has already saved it, so the address now
  carries the app version — filled in by the running app, so it can never go stale. Every release
  makes your phone fetch the icon again.

**To see it:** delete the old shortcut from your home screen, then add it again. The tab icon is
unchanged — it is inside the browser, where rounded corners were always correct.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
