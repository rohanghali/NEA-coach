#!/bin/bash
#
#  NEA Coach — double-click this file to start.
#
#  What it does:
#    1. Finds Node.js. If it is missing, downloads the correct build for this Mac
#       (Intel or Apple Silicon) into the app folder. No installer, no admin password.
#    2. Starts the NEA Coach server.
#    3. Opens your browser to it.
#
#  It never modifies your PowerPoint file. See README.md.
#
#  To stop the app: close this window, or press Control-C in it.

set -u

# Always work relative to wherever this script actually lives, so it works
# from Dropbox, the Desktop, or anywhere else.
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$APP_DIR" || exit 1

PORT="${NEA_COACH_PORT:-4173}"
RUNTIME_DIR="$APP_DIR/runtime"
NODE_BIN=""

# ---------------------------------------------------------------- pretty output
BOLD="$(printf '\033[1m')"; DIM="$(printf '\033[2m')"
GREEN="$(printf '\033[32m')"; YELLOW="$(printf '\033[33m')"
RED="$(printf '\033[31m')"; RESET="$(printf '\033[0m')"

say()  { printf '%s\n' "$1"; }
ok()   { printf '%s  ✓%s %s\n' "$GREEN" "$RESET" "$1"; }
warn() { printf '%s  !%s %s\n' "$YELLOW" "$RESET" "$1"; }
bad()  { printf '%s  ✗%s %s\n' "$RED" "$RESET" "$1"; }

say ""
say "${BOLD}  NEA Coach${RESET} ${DIM}— portfolio reviewer${RESET}"
say "${DIM}  ------------------------------------------------------------${RESET}"
say ""

# ---------------------------------------------------------------- 1. find Node
find_node() {
  # a) An explicitly configured interpreter wins.
  if [ -n "${NEA_COACH_NODE:-}" ] && [ -x "${NEA_COACH_NODE}" ]; then
    NODE_BIN="${NEA_COACH_NODE}"; return 0
  fi

  # b) A previously downloaded runtime inside the app folder.
  if [ -x "$RUNTIME_DIR/node" ]; then
    NODE_BIN="$RUNTIME_DIR/node"; return 0
  fi

  # c) A system-wide Node, if the student already has one.
  if command -v node >/dev/null 2>&1; then
    NODE_BIN="$(command -v node)"
    # Require a version new enough to run the app (needs 18+).
    local major
    major="$("$NODE_BIN" -p 'process.versions.node.split(".")[0]' 2>/dev/null || echo 0)"
    if [ "${major:-0}" -ge 18 ] 2>/dev/null; then return 0; fi
    warn "Found Node $major, which is too old. A newer one will be downloaded instead."
    NODE_BIN=""
  fi

  return 1
}

download_node() {
  # Detect the chip so we fetch the right build.
  local arch
  arch="$(uname -m)"
  case "$arch" in
    arm64)  NODE_ARCH="arm64" ;;
    x86_64) NODE_ARCH="x64" ;;
    *)
      bad "This Mac reports an unrecognised processor type: $arch"
      say "     Please tell whoever set this up, quoting that value."
      return 1
      ;;
  esac

  # Node 22 is the long-term-support line and still supports macOS 11+,
  # which matters because this Mac's exact age is unknown.
  local version="v22.14.0"
  local tarball="node-${version}-darwin-${NODE_ARCH}.tar.gz"
  local base="https://nodejs.org/dist/${version}"
  local tmp
  tmp="$(mktemp -d)"
  local url="${base}/${tarball}"

  say "  Setting up for the first time (${NODE_ARCH} Mac)."
  say "  ${DIM}Downloading the JavaScript engine, about 45 MB. One time only.${RESET}"
  say ""

  if command -v curl >/dev/null 2>&1; then
    if ! curl -fL --progress-bar --connect-timeout 20 --max-time 900 -o "$tmp/$tarball" "$url"; then
      bad "The download failed. Check the internet connection and try again."
      say "     ${DIM}$url${RESET}"
      rm -rf "$tmp"
      return 1
    fi
  else
    bad "curl is missing, so the download cannot run. That is unusual on a Mac."
    rm -rf "$tmp"
    return 1
  fi

  say ""
  say "  Unpacking..."
  if ! tar -xzf "$tmp/$tarball" -C "$tmp"; then
    bad "The download could not be unpacked. It may have been incomplete."
    rm -rf "$tmp"
    return 1
  fi

  mkdir -p "$RUNTIME_DIR"
  # Copy just the binary — the app needs nothing else from the distribution.
  if ! cp "$tmp/node-${version}-darwin-${NODE_ARCH}/bin/node" "$RUNTIME_DIR/node"; then
    bad "Could not install the engine into the app folder."
    rm -rf "$tmp"
    return 1
  fi
  chmod +x "$RUNTIME_DIR/node"

  # macOS quarantine flag: clear it so the binary can run.
  xattr -d com.apple.quarantine "$RUNTIME_DIR/node" 2>/dev/null || true

  # Check it actually runs before trusting it.
  if ! "$RUNTIME_DIR/node" -e 'process.exit(0)' 2>/dev/null; then
    bad "The downloaded engine will not run on this Mac."
    say "     ${DIM}This can happen with very old versions of macOS.${RESET}"
    rm -rf "$tmp"
    return 1
  fi

  rm -rf "$tmp"
  NODE_BIN="$RUNTIME_DIR/node"
  return 0
}

if find_node; then
  ok "Ready to go."
else
  if ! download_node; then
    say ""
    say "  ${BOLD}Setup could not finish.${RESET}"
    say "  Nothing has been changed or broken. Please send this message on:"
    say "  ${DIM}platform=$(uname -s) $(uname -r) arch=$(uname -m)${RESET}"
    say ""
    read -r -p "  Press Return to close this window. " _ || true
    exit 1
  fi
  ok "Setup complete."
fi

# ---------------------------------------------------------------- 2. sanity check
if ! "$NODE_BIN" --version >/dev/null 2>&1; then
  bad "The JavaScript engine found at $NODE_BIN will not run."
  read -r -p "  Press Return to close this window. " _ || true
  exit 1
fi

if ! "$NODE_BIN" --check "$APP_DIR/server.js" >/dev/null 2>&1; then
  bad "The application files look damaged (server.js will not parse)."
  say "     ${DIM}If Dropbox is still syncing, wait for it to finish and try again.${RESET}"
  read -r -p "  Press Return to close this window. " _ || true
  exit 1
fi

# ---------------------------------------------------------------- 3. already running?
if command -v curl >/dev/null 2>&1; then
  if curl -fsS --max-time 2 "http://127.0.0.1:${PORT}/api/health" >/dev/null 2>&1; then
    ok "NEA Coach is already running."
    say ""
    open "http://127.0.0.1:${PORT}" 2>/dev/null || true
    say "  Opening your browser. This window can be closed."
    say ""
    exit 0
  fi
fi

# ---------------------------------------------------------------- 4. start it
say ""
say "  ${DIM}Starting. A browser window will open in a moment.${RESET}"
say "  ${DIM}Leave this window open while you use the app. Control-C quits.${RESET}"
say ""

# Open the browser once the server is actually listening, rather than after a
# fixed delay, so the student never sees a "cannot connect" page.
(
  for _ in $(seq 1 60); do
    if command -v curl >/dev/null 2>&1; then
      if curl -fsS --max-time 1 "http://127.0.0.1:${PORT}/api/health" >/dev/null 2>&1; then
        open "http://127.0.0.1:${PORT}" 2>/dev/null || true
        exit 0
      fi
    else
      sleep 1.5
      open "http://127.0.0.1:${PORT}" 2>/dev/null || true
      exit 0
    fi
    sleep 0.5
  done
) &

NEA_COACH_PORT="$PORT" "$NODE_BIN" "$APP_DIR/server.js"
STATUS=$?

# ---------------------------------------------------------------- 5. exit
say ""
if [ "$STATUS" -ne 0 ]; then
  bad "NEA Coach stopped with an error (code $STATUS)."
  say "     ${DIM}The lines above this message explain what happened.${RESET}"
else
  ok "NEA Coach has stopped."
fi
say ""
read -r -p "  Press Return to close this window. " _ || true
