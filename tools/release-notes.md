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

Nothing about how the app works has changed. This one is about the tools that build and publish
it — the path between the code being right and you receiving it.

- **A release can no longer be published half-uploaded.** It is now created as a draft, both
  packages are attached, and only then is it published. Before, the release went public first, so
  a failure part-way through uploading left a live release missing one of its two downloads — and
  because it was the newest, the download link served exactly that. A failure now removes the
  incomplete release instead of leaving it, and re-running works.
- **A Windows package can no longer be built without Node in it.** The build used to drop a note
  in the runtime folder and carry on, so the package could promise "Nothing to install. Node.js is
  already inside this folder" while containing no Node at all. It now refuses to build.
- **A typo in the build command no longer deletes the packages.** An unrecognised argument wiped
  every built zip and *then* printed the usage message.

**If you are updating:** quit NEA Coach completely first, then open the new copy.

**In v1.0.42**

- Automatic re-checking works again (two separate silent faults); the Mac instructions no longer
  mention a Windows file or a frozen version; and eight smaller corrections.

**If you have an older copy already:** quit it first (close the black "NEA Coach server"
window) before starting this one, or the old version will keep answering your browser.

It never modifies your original PowerPoint, and nothing leaves your computer except the
pages a check sends to Google.
