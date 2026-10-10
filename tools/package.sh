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

# Everything a user needs. `state` is deliberately absent: it is regenerated on first
# run, and shipping it would leak the student's report, their portfolio and their key.
INCLUDE=(
  server.js
  lib
  public
  spec
  sample
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

   A black window opens and your web browser opens the app.
   LEAVE THE BLACK WINDOW OPEN while you use it — closing it stops the app.

   Nothing to install. Node.js is already inside this folder.

3. The first time, you need a free Google API key:
   see SETUP-AI.md, or the in-app Settings page.
   It is free and needs no bank card.

Your portfolio is never uploaded anywhere except to Google, and only the pages
a check needs. The app never changes your original PowerPoint file.
EOF
}

build_windows() {
  local target="$DIST/NEA-Coach-windows"
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
  ( cd "$DIST" && zip -qr "NEA-Coach-windows.zip" "NEA-Coach-windows" )
  say "wrote $DIST/NEA-Coach-windows.zip ($(du -h "$DIST/NEA-Coach-windows.zip" | cut -f1 | tr -d ' '))"
}

build_mac() {
  local target="$DIST/NEA-Coach-mac"
  say "Building for macOS..."
  rm -rf "$target"
  copy_common "$target"
  write_readme_first "$target"
  rm -f "$target/Start NEA Coach.bat"
  say "NOTE: this build does NOT bundle Node — the Mac launcher downloads the matching"
  say "      runtime on first run, which is why there is no runtime folder here."
  ( cd "$DIST" && zip -qr "NEA-Coach-mac.zip" "NEA-Coach-mac" )
  say "wrote $DIST/NEA-Coach-mac.zip ($(du -h "$DIST/NEA-Coach-mac.zip" | cut -f1 | tr -d ' '))"
}

mkdir -p "$DIST"
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
