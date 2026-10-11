#!/bin/bash
#
# Publish a release to GitHub, with the download attached.
#
#   ./tools/release.sh 1.0.0
#
# THE TOKEN IS NEVER PRINTED. It is read from the macOS keychain, where the first
# `git push` put it, and passed to curl through a config file rather than the command
# line — a token on a command line is visible to anything that can list processes, and
# curl prints the URL it is using in some errors.
#
# GitHub's website caps uploads at 25 MB and the Windows package is about 31 MB, which
# is why this exists at all: the release has to be made through the API.

set -u

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="${1:-}"

say() { printf '  %s\n' "$1"; }

if [ -z "$VERSION" ]; then
  echo "  usage: tools/release.sh <version>       e.g. tools/release.sh 1.0.0"
  exit 1
fi

# Tolerate either "1.0.0" or "v1.0.0", and always create the tag as vX.Y.Z.
TAG="$VERSION"
case "$TAG" in v*) ;; *) TAG="v$VERSION" ;; esac

cd "$APP_DIR"

REPO=$(git remote get-url origin 2>/dev/null | sed -E 's#.*github\.com[:/]([^/]+/[^/.]+)(\.git)?#\1#')
if [ -z "$REPO" ]; then
  say "Could not work out the GitHub repository from 'git remote -v'."
  exit 1
fi
say "Repository: $REPO"

# --- the token, from the keychain, without echoing it -------------------------------
TOKEN=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill 2>/dev/null | sed -n 's/^password=//p')
if [ -z "$TOKEN" ]; then
  say "No GitHub token found in the keychain."
  say "Push once with 'git push' and enter your token as the password, then try again."
  exit 1
fi
say "Token: found in the keychain (not shown)"

# Keep it out of the process list and out of logs.
AUTH_FILE="$(mktemp)"
trap 'rm -f "$AUTH_FILE" "${RESP_FILE:-}" 2>/dev/null' EXIT
chmod 600 "$AUTH_FILE"
printf 'header = "Authorization: Bearer %s"\nheader = "Accept: application/vnd.github+json"\n' "$TOKEN" > "$AUTH_FILE"
unset TOKEN

# --- the package --------------------------------------------------------------------
ZIP="dist/NEA-Coach-windows.zip"
if [ ! -f "$ZIP" ]; then
  say "No package found. Building it first..."
  ./tools/package.sh windows || exit 1
fi
SIZE_MB=$(( $(wc -c < "$ZIP") / 1048576 ))
say "Package: $ZIP (${SIZE_MB} MB)"

# --- check the tag is free ----------------------------------------------------------
RESP_FILE="$(mktemp)"
CODE=$(curl -s -o "$RESP_FILE" -w '%{http_code}' --max-time 20 \
  --config "$AUTH_FILE" "https://api.github.com/repos/$REPO/releases/tags/$TAG")
if [ "$CODE" = "200" ]; then
  say "A release for $TAG already exists. Use a different version, or delete it on GitHub."
  exit 1
fi

# --- create the release -------------------------------------------------------------
BODY='Reviews a D&T NEA portfolio and tells you what to fix, in order of what is costing you marks.

**To use it:**

1. Download `NEA-Coach-windows.zip` below and unzip it.
2. Double-click `Start NEA Coach.bat`. **Nothing needs installing** — Node.js is inside the download.
3. Get a free Google API key (no card needed) — the app explains it, and `SETUP-AI.md` is in the folder.

On a Mac, the repo itself runs as-is: double-click `Start NEA Coach.command`.

It never modifies your original PowerPoint, and nothing leaves your computer except the pages a check sends to Google.'

say "Creating release $TAG..."
CODE=$(curl -s -o "$RESP_FILE" -w '%{http_code}' --max-time 60 \
  --config "$AUTH_FILE" -X POST \
  "https://api.github.com/repos/$REPO/releases" \
  -d "$(printf '{"tag_name":"%s","name":"NEA Coach %s","body":%s,"draft":false,"prerelease":false}' \
        "$TAG" "$VERSION" "$(printf '%s' "$BODY" | "${APP_DIR}/runtime/node" -e 'let s="";process.stdin.on("data",d=>s+=d).on("end",()=>process.stdout.write(JSON.stringify(s)))')")")

if [ "$CODE" != "201" ]; then
  say "GitHub refused to create the release (HTTP $CODE)."
  "${APP_DIR}/runtime/node" -e '
    try { const j=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));
      if (j.message) console.log("  " + j.message);
      if (j.errors) JSON.stringify(j.errors).slice(0,300) && console.log("  " + JSON.stringify(j.errors).slice(0,300));
    } catch (e) {}' "$RESP_FILE"
  exit 1
fi

RELEASE_ID=$("${APP_DIR}/runtime/node" -e '
  const j=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));
  process.stdout.write(String(j.id));' "$RESP_FILE")
say "Release created (id $RELEASE_ID). Uploading the package..."

# --- upload the zip -----------------------------------------------------------------
CODE=$(curl -s -o "$RESP_FILE" -w '%{http_code}' --max-time 900 \
  --config "$AUTH_FILE" -X POST \
  -H "Content-Type: application/zip" \
  --data-binary "@$ZIP" \
  "https://uploads.github.com/repos/$REPO/releases/$RELEASE_ID/assets?name=NEA-Coach-windows.zip")

if [ "$CODE" = "201" ]; then
  say "Uploaded."
else
  say "Upload failed (HTTP $CODE). The release exists but has no file attached."
  "${APP_DIR}/runtime/node" -e '
    try { const j=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));
      if (j.message) console.log("  " + j.message); } catch (e) {}' "$RESP_FILE"
  exit 1
fi

echo
say "Done. The download link to share is:"
say "  https://github.com/$REPO/releases/latest"
