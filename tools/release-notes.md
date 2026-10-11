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

- **Automatic re-checking actually happens now.** Two faults, both silent. If you edited a page
  while a check was running, the change was thrown away — and a check takes minutes, so the report
  you were looking at described the version *before* your edit with no sign anything was wrong. It
  is remembered and re-checked when the scan finishes. And if you changed any setting within a few
  seconds of saving your work, that cancelled the pending re-check, because saving a setting
  restarts the file watcher and the restart mistook your edit for "nothing changed".
- **The Mac download no longer tells you to double-click a Windows file.** The Mac package shipped
  the Windows instructions at the top — "Double-click Start NEA Coach.bat" — for a file that is not
  in the zip, so the first thing a Mac user read was a File-not-found.
- **The Gatekeeper command in the Mac instructions works again.** It pointed at a folder named
  after version 1.0.29, so anyone who copied it got an app macOS still refused to open, with no
  error to explain why. It no longer names a version at all.

**Smaller corrections in this version**

- The whole-portfolio cross-check request is now counted against your daily allowance; it was not,
  so every completed check under-reported by one.
- A check that falls back to a backup model now prices the run at the model that actually answered,
  not the one that failed.
- Switching deep thinking on now invalidates the page cache, so the setting does something.
- The cost hint before a check can finally say a repeat check is free — the page was looking for a
  list where the app sends a number, so it always quoted a full re-read.
- The exported report puts mark-costing findings first; "substance" has rank 0, which the sort
  treated as "unranked".

**In v1.0.41**

- Cancel no longer discards the pages already read; the daily-allowance percentage can finally
  work; a failed page read is no longer reported as a successful review; two dark-mode contrast
  fixes. **The test suite no longer reads your real data folder.**

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
