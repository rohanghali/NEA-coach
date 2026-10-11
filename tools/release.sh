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

# --- the packages -------------------------------------------------------------------
# BOTH platforms ship with every release. Shipping only the Windows zip meant a Mac
# user had no launcher to double-click and nothing to download, which is exactly the
# question this answers.
ASSETS=(dist/NEA-Coach-windows.zip dist/NEA-Coach-mac.zip)

# The packages are ALWAYS rebuilt, and never reused just because dist/ is not empty.
#
# They used to be built only when the zips were missing. Releasing v1.0.22 while
# v1.0.21 zips were still sitting in dist/ therefore uploaded v1.0.21 to the v1.0.22
# release. Nothing failed and the log looked clean - the only hint was a "0 MB" line
# that was itself a bug (integer division of a 343 KB file). Worse, the launcher
# compares its own version against the server it starts, so a wrong-but-consistent
# package reports no mismatch and the student silently gets the old build. That is
# the exact failure the versioned folders were introduced to kill.
say "Building both packages from the current source..."
rm -f "${ASSETS[@]}"
./tools/package.sh both || exit 1

APP_VERSION="$(sed -n "s/^const VERSION = '\([0-9.]*\)'.*/\1/p" server.js | head -1)"
EXPECT_VERSION="${TAG#v}"
for z in "${ASSETS[@]}"; do
  [ -f "$z" ] || { say "Package $z is still missing after building."; exit 1; }
  say "Package: $z ($(( $(wc -c < "$z") / 1024 )) KB)"
done

# The version is baked into the folder name inside each zip, and that folder name is
# the one thing a stale build cannot fake. If it does not match the tag being
# released, the release does not go out.
for z in "${ASSETS[@]}"; do
  ROOT_DIR="$(unzip -Z1 "$z" 2>/dev/null | head -1)"
  ROOT_DIR="${ROOT_DIR%%/*}"
  case "$ROOT_DIR" in
    *-"$EXPECT_VERSION") ;;
    *) say "REFUSING TO RELEASE: $z contains '$ROOT_DIR', which is not a $EXPECT_VERSION package."
       say "  server.js currently says $APP_VERSION."
       exit 1 ;;
  esac
done
say "Both packages contain a $EXPECT_VERSION build."

# --- check the tag is free ----------------------------------------------------------
RESP_FILE="$(mktemp)"
CODE=$(curl -s -o "$RESP_FILE" -w '%{http_code}' --max-time 20 \
  --config "$AUTH_FILE" "https://api.github.com/repos/$REPO/releases/tags/$TAG")
if [ "$CODE" = "200" ]; then
  say "A release for $TAG already exists. Use a different version, or delete it on GitHub."
  exit 1
fi

# --- create the release -------------------------------------------------------------
# The notes live in tools/release-notes.md and are read in, rather than being a
# BODY='...' string inside this script.
#
# They used to be single-quoted here, and they are prose: they contain an apostrophe
# ("the AI's own reading") and backticks (for folder names). The apostrophe closed the
# string early, the rest was executed as shell, and BODY ended up unset - which, under
# set -u, produced a release with no notes at all and no error worth noticing.
BODY="$(cat "$APP_DIR/tools/release-notes.md")"

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
say "Release created (id $RELEASE_ID). Uploading the packages..."

# --- upload every zip ---------------------------------------------------------------
for ZIP in "${ASSETS[@]}"; do
  NAME="$(basename "$ZIP")"
  say "Uploading $NAME..."
  CODE=$(curl -s -o "$RESP_FILE" -w '%{http_code}' --max-time 900 \
    --config "$AUTH_FILE" -X POST \
    -H "Content-Type: application/zip" \
    --data-binary "@$ZIP" \
    "https://uploads.github.com/repos/$REPO/releases/$RELEASE_ID/assets?name=$NAME")

  if [ "$CODE" = "201" ]; then
    say "  Uploaded $NAME."
  else
    say "  Upload failed for $NAME (HTTP $CODE)."
    "${APP_DIR}/runtime/node" -e '
      try { const j=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));
        if (j.message) console.log("  " + j.message); } catch (e) {}' "$RESP_FILE"
    exit 1
  fi
done

echo
say "Done. The download link to share is:"
say "  https://github.com/$REPO/releases/latest"
