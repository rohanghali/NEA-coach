#!/bin/bash
#
# Download the latest NEA Coach release.
#
#   ./tools/get-release.sh              the Mac build, into this folder
#   ./tools/get-release.sh windows      the Windows build
#   ./tools/get-release.sh mac 1.0.30   a specific version
#
# WHY THIS EXISTS
#
# The one-line curl command that used to be given out stopped working the moment the
# repository was made private, and the reason is not obvious: on a private repository the
# ordinary release URL
#
#   https://github.com/OWNER/REPO/releases/download/TAG/FILE
#
# returns HTTP 404 EVEN WITH A TOKEN SUPPLIED. Private release files are only ever served
# through the GitHub API, at /repos/.../releases/assets/<id>. So "just add your token" does
# not fix it, which is exactly the sort of thing that costs an hour to work out.
#
# This tries the plain public URL first — which needs no token and is the honest way for a
# public repository — and falls back to the API, using the token git already keeps in the
# keychain, only when that fails.
#
# The script is deliberately a plain shell script with no dependencies beyond curl and git,
# because it has to work on a machine where NEA Coach has never been run.

set -u

REPO="rohanghali/NEA-coach"
WANT="${1:-mac}"
VERSION="${2:-}"

case "$WANT" in
  mac|windows) ;;
  *) echo "  usage: tools/get-release.sh [mac|windows] [version]"; exit 1 ;;
esac

FILE="NEA-Coach-$WANT.zip"
say() { printf '  %s\n' "$1"; }

if [ -n "$VERSION" ]; then
  API="https://api.github.com/repos/$REPO/releases/tags/v${VERSION#v}"
  WHERE="v${VERSION#v}"
else
  API="https://api.github.com/repos/$REPO/releases/latest"
  WHERE="the latest release"
fi

# --- what is already here -----------------------------------------------------------
if [ -f "$FILE" ]; then
  say "note: $FILE already exists here and will be replaced"
fi

# --- the ordinary way, which needs no token on a public repository -------------------
say "Looking up $WHERE..."
PLAIN=$(curl -sS -H "Accept: application/vnd.github+json" "$API" 2>/dev/null \
  | sed -n 's/.*"browser_download_url": *"\([^"]*\/'"$FILE"'\)".*/\1/p' | head -1)

if [ -n "$PLAIN" ]; then
  say "Downloading $FILE..."
  if curl -fL --progress-bar -o "$FILE" "$PLAIN" 2>/dev/null; then
    say "Saved $(pwd)/$FILE ($(wc -c < "$FILE" | tr -d ' ') bytes)"
    exit 0
  fi
  say "The public link did not work — this repository is probably private."
  rm -f "$FILE"
fi

# --- the API way, which is the only one a private repository serves ------------------
say "Asking GitHub for the file directly..."
TOKEN=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill 2>/dev/null | sed -n 's/^password=//p')

if [ -z "$TOKEN" ]; then
  say "This repository is private and no GitHub token is stored in the keychain, so there"
  say "is nothing to authenticate with. Either make the repository public, or push once"
  say "with 'git push' and give your token as the password, then run this again."
  exit 1
fi

META=$(curl -sS -H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github+json" "$API" 2>/dev/null)
if [ -z "$META" ]; then
  say "GitHub did not answer. Check this Mac is online."
  exit 1
fi

# The asset's own "id" line comes just before its "name" line in the reply, so the last id
# seen before the matching name is the one to fetch.
ASSET=$(printf '%s\n' "$META" | awk -v want="\"name\": \"$FILE\"" '
  /"id":/ { id = $2; gsub(/,/, "", id) }
  index($0, want) { print id; exit }
')

if [ -z "$ASSET" ]; then
  say "The release was found but it has no $FILE attached."
  say "Files in it:"
  printf '%s\n' "$META" | sed -n 's/.*"name": "\(NEA-Coach[^"]*\)".*/    \1/p'
  exit 1
fi

say "Downloading $FILE through the API..."
if ! curl -fL --progress-bar -o "$FILE" \
  -H "Authorization: Bearer $TOKEN" \
  -H "Accept: application/octet-stream" \
  "https://api.github.com/repos/$REPO/releases/assets/$ASSET"; then
  say "The download failed. The token may have expired, or not have access to this repository."
  rm -f "$FILE"
  exit 1
fi

# A private asset that is refused comes back as a small JSON error saved as a .zip, so the
# size is checked rather than trusted: an HTML or JSON error page is not a zip and would
# otherwise be unzipped into nothing later with a confusing message.
if ! unzip -tq "$FILE" >/dev/null 2>&1; then
  say "What came back is not a zip file, so the download did not really succeed."
  say "It starts with: $(head -c 120 "$FILE" | tr -d '\n')"
  rm -f "$FILE"
  exit 1
fi

say "Saved $(pwd)/$FILE ($(wc -c < "$FILE" | tr -d ' ') bytes)"
say "Unzip it, then open the app inside."
