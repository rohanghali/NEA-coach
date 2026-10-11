#!/bin/bash
#
# Build the shareable package.
#
#   ./tools/package.sh            build for Windows and Mac
#   ./tools/package.sh windows    Windows only
#   ./tools/package.sh mac        Mac only
#
# WHY THIS EXISTS
#
# The working folder contains 556 MB of the student's OWN work in `state/` — reports,
# cached pages, a copy of their portfolio and their API key — and a 104 MB Node runtime
# that only runs on the machine it was built for. Neither belongs in something sent to
# someone else. This assembles exactly what a new user needs and nothing else.
#
# The output is a folder anyone can unzip and run, with Node already inside it, so the
# recipient installs NOTHING.

set -u

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
DIST="$APP_DIR/dist"
WHICH="${1:-both}"
NODE_VERSION="v22.14.0"
# The version the app reports, used in the folder names. Folders MUST be versioned:
# every download used to unzip to "NEA-Coach-mac", so a second one landed beside the
# first as "NEA-Coach-mac 2", then " 3", " 4"... and the older copy kept the port.
# Versioned folders cannot collide, so one app always means one folder.
APP_VERSION="$(sed -n "s/^const VERSION = '\([0-9.]*\)'.*/\1/p" "$APP_DIR/server.js" | head -1)"
[ -n "$APP_VERSION" ] || { echo "  Could not read VERSION from server.js"; exit 1; }

# Everything a user needs. `state` is deliberately absent: it is regenerated on first
# run, and shipping it would leak the student's report, their portfolio and their key.
INCLUDE=(
  server.js
  lib
  public
  spec
  sample
  assets
  tools
  README.md
  SETUP-AI.md
  WISHLIST.txt
  "Start NEA Coach.bat"
  "Start NEA Coach.command"
)

say() { printf '  %s\n' "$1"; }

# Copy the shared files into a target folder.
copy_common() {
  local target="$1"
  mkdir -p "$target"
  for item in "${INCLUDE[@]}"; do
    if [ -e "$APP_DIR/$item" ]; then
      cp -R "$APP_DIR/$item" "$target/"
    else
      say "note: $item is missing, skipped"
    fi
  done
  # Never ship these, even if a stray copy exists.
  rm -rf "$target/state" "$target/dist" 2>/dev/null
  find "$target" -name '.DS_Store' -delete 2>/dev/null
}

# A short, obvious first page for someone who has just unzipped it.
write_readme_first() {
  local target="$1"
  cat > "$target/START HERE.txt" <<'EOF'
NEA Coach
=========

1. Unzip this folder somewhere you can find it, such as your Documents.

2. Double-click "Start NEA Coach.bat"

   A shortcut called "NEA Coach" with its own icon appears on your Desktop the
   first time this runs, so afterwards you can start it from there instead.

   Two black windows open. One is titled "NEA Coach server" — that IS the app,
   LEAVE IT OPEN while you use it; closing it stops the app. The other window
   just waits for the server and then tells you the address; it can be closed.

   Nothing to install. Node.js is already inside this folder.

3. The first time, you need a free Google API key:
   see SETUP-AI.md, or the in-app Settings page.
   It is free and needs no bank card.

Your portfolio is never uploaded anywhere except to Google, and only the pages
a check needs. The app never changes your original PowerPoint file.
EOF
}

# Mac users meet Gatekeeper because the app is not signed with a paid Apple
# certificate. Warn them BEFORE the dialog appears, or they click "Move to Bin"
# and conclude the app is broken.
write_readme_first_mac() {
  local target="$1"
  cat >> "$target/START HERE.txt" <<'EOF'

--------------------------------------------------------------

TO PUT IT ON YOUR DESKTOP OR IN APPLICATIONS

This folder contains "NEA Coach.app". Drag it wherever you like - the Desktop,
or Applications. It is a real Mac app with its own icon, and it carries
everything it needs inside it, so it keeps working after being moved.

Double-click it and the same server window opens as with the file below.

--------------------------------------------------------------

If macOS says it cannot verify "NEA Coach.app" or "Start NEA Coach.command":

That is Gatekeeper. It means the file came from the internet and is not signed
with a paid Apple certificate - NOT that anything harmful was found.

Fix it once, either way:

  a) Right-click (or Control-click) the app, choose Open, then click Open again
     in the box that appears. Do this once and macOS remembers.

  b) If that does not work, open Terminal (Command-Space, type Terminal) and
     run this line, then double-click the file again:

       xattr -dr com.apple.quarantine ~/Downloads/NEA-Coach-mac-1.0.29

     Change the path if you unzipped it somewhere other than Downloads.

You may need to do this again after downloading a new version.

The first launch also downloads Node.js automatically (about 40 MB) and shows
a progress bar. Every launch after that is instant.
EOF
}

build_windows() {
  local target="$DIST/NEA-Coach-windows-$APP_VERSION"
  say "Building for Windows..."
  rm -rf "$target"
  copy_common "$target"
  write_readme_first "$target"

  mkdir -p "$target/runtime"
  if [ -f "$APP_DIR/runtime/node.exe" ]; then
    cp "$APP_DIR/runtime/node.exe" "$target/runtime/node.exe"
    say "bundled runtime/node.exe ($(du -h "$target/runtime/node.exe" | cut -f1 | tr -d ' '))"
  else
    cat > "$target/runtime/PUT-NODE-EXE-HERE.txt" <<'EOF'
This folder should contain node.exe, the Windows copy of Node.js.

It is missing from this build. Either:

  a) run tools/package.sh windows on a machine with internet access, which fetches
     it automatically, or

  b) download the Windows .zip from https://nodejs.org, open it, and copy
     node.exe into this runtime folder.

Windows users otherwise need to install Node.js from https://nodejs.org themselves.
EOF
    say "WARNING: runtime/node.exe not present — see the note placed in runtime/"
  fi

  # The Mac launcher is noise on Windows.
  rm -f "$target/Start NEA Coach.command"

  # .bat files MUST have CRLF line endings, and an editor on macOS or Linux will
  # happily strip them. cmd.exe can mis-parse a bare-LF batch file, and it would fail
  # on the recipient's machine rather than here, which is the worst place for it to
  # surface. Converted at package time so it cannot regress.
  local bat="$target/Start NEA Coach.bat"
  if [ -f "$bat" ]; then
    perl -pi -e 's/\r?\n/\r\n/' "$bat"
    say "converted Start NEA Coach.bat to CRLF line endings"
  fi
  ( cd "$DIST" && zip -qr "NEA-Coach-windows.zip" "NEA-Coach-windows-$APP_VERSION" )
  say "wrote $DIST/NEA-Coach-windows.zip ($(du -h "$DIST/NEA-Coach-windows.zip" | cut -f1 | tr -d ' '))"
}

# A real Mac app, so the icon can sit on the Desktop or in Applications.
#
# The bundle carries its OWN copy of the app in Contents/Resources/app rather than pointing
# at the folder beside it. That is the difference between an app you can drag to the Desktop
# and a shortcut that breaks the moment it is moved — and being able to put it on the
# Desktop is the entire reason it exists.
#
# It is staged OUTSIDE the target and moved in at the end. Copying the target into a folder
# inside the target would recurse into the copy being written.
build_mac_app() {
  local target="$1"
  local stage="$DIST/.nea-app-stage"
  local bundle="$target/NEA Coach.app"

  rm -rf "$stage" "$bundle"
  mkdir -p "$stage/Contents/MacOS" "$stage/Contents/Resources"

  cp -R "$target" "$stage/Contents/Resources/app"
  # Belt and braces: nothing stripped, personal or duplicated may survive inside a bundle
  # that someone might email on.
  rm -rf "$stage/Contents/Resources/app/state" \
         "$stage/Contents/Resources/app/dist" \
         "$stage/Contents/Resources/app/NEA Coach.app" 2>/dev/null

  cp "$APP_DIR/assets/NEA Coach.icns" "$stage/Contents/Resources/NEA Coach.icns"

  cat > "$stage/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>NEA Coach</string>
  <key>CFBundleDisplayName</key><string>NEA Coach</string>
  <key>CFBundleIdentifier</key><string>uk.co.neacoach.app</string>
  <key>CFBundleExecutable</key><string>NEA Coach</string>
  <key>CFBundleIconFile</key><string>NEA Coach</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleSignature</key><string>????</string>
  <key>CFBundleInfoDictionaryVersion</key><string>6.0</string>
  <key>CFBundleShortVersionString</key><string>$APP_VERSION</string>
  <key>CFBundleVersion</key><string>$APP_VERSION</string>
  <key>LSMinimumSystemVersion</key><string>10.15</string>
  <key>NSHighResolutionCapable</key><true/>
  <key>LSApplicationCategoryType</key><string>public.app-category.education</string>
</dict>
</plist>
PLIST

  # The entry point. It does not run the server itself: the server prints the address your
  # phone should use and says how to stop it, and that needs a window. So it opens the same
  # launcher you would double-click, in a Terminal window, and steps out of the way.
  cat > "$stage/Contents/MacOS/NEA Coach" <<'LAUNCH'
#!/bin/bash
set -u
HERE="$(cd "$(dirname "$0")" && pwd)"
LAUNCHER="$HERE/../Resources/app/Start NEA Coach.command"
if [ ! -f "$LAUNCHER" ]; then
  osascript -e 'display alert "NEA Coach is incomplete" message "The app inside this bundle is missing. Download it again from the NEA Coach releases page." as critical' >/dev/null 2>&1
  exit 1
fi
chmod +x "$LAUNCHER" 2>/dev/null
exec open -a Terminal "$LAUNCHER"
LAUNCH
  chmod +x "$stage/Contents/MacOS/NEA Coach"
  # The launcher has to keep its executable bit through cp and through zip.
  chmod +x "$stage/Contents/Resources/app/Start NEA Coach.command" 2>/dev/null

  mv "$stage" "$bundle"
  say "built NEA Coach.app (drag it to the Desktop or Applications)"
}

build_mac() {
  local target="$DIST/NEA-Coach-mac-$APP_VERSION"
  say "Building for macOS..."
  rm -rf "$target"
  copy_common "$target"
  write_readme_first "$target"
  write_readme_first_mac "$target"
  rm -f "$target/Start NEA Coach.bat"
  build_mac_app "$target"
  say "NOTE: this build does NOT bundle Node — the Mac launcher downloads the matching"
  say "      runtime on first run, which is why there is no runtime folder here."
  ( cd "$DIST" && zip -qr "NEA-Coach-mac.zip" "NEA-Coach-mac-$APP_VERSION" )
  say "wrote $DIST/NEA-Coach-mac.zip ($(du -h "$DIST/NEA-Coach-mac.zip" | cut -f1 | tr -d ' '))"
}

mkdir -p "$DIST"
# Clear unpacked folders from earlier builds ONCE, before either platform runs.
# Leftovers were being swept into the zip by the glob, shipping an archive that
# contained two app folders. This must NOT live in copy_common: that runs per
# platform, and would delete the other platform's finished output.
rm -rf "$DIST"/NEA-Coach-mac* "$DIST"/NEA-Coach-windows* "$DIST"/.nea-app-stage 2>/dev/null
say "App folder: $APP_DIR"
echo
case "$WHICH" in
  windows) build_windows ;;
  mac) build_mac ;;
  both) build_windows; echo; build_mac ;;
  *) echo "  usage: tools/package.sh [windows|mac|both]"; exit 1 ;;
esac
echo
say "Done. Nothing in dist/ contains state/, so no report, portfolio or API key is included."
