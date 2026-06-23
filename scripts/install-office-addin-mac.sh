#!/bin/sh

set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
REPO_DIR=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)

if [ "${1:-}" != "" ]; then
  MANIFEST_PATH=$1
elif [ -f "$REPO_DIR/dist/manifest.xml" ]; then
  MANIFEST_PATH="$REPO_DIR/dist/manifest.xml"
else
  MANIFEST_PATH="$REPO_DIR/manifest.xml"
fi

if [ ! -f "$MANIFEST_PATH" ]; then
  printf 'Manifest file not found: %s\n' "$MANIFEST_PATH" >&2
  exit 1
fi

resolve_path() {
  input_path=$1

  input_dir=$(CDPATH= cd -- "$(dirname -- "$input_path")" && pwd)
  printf '%s/%s\n' "$input_dir" "$(basename -- "$input_path")"
}

MANIFEST_ABS=$(resolve_path "$MANIFEST_PATH")
ADDIN_ID=$(perl -0ne 'print $1 if /<Id>\s*([^<]+)\s*<\/Id>/s' "$MANIFEST_ABS")
SOURCE_LOCATION=$(perl -0ne 'print $1 if /<SourceLocation[^>]*DefaultValue="([^"]+)"/s' "$MANIFEST_ABS")

if [ -z "$ADDIN_ID" ]; then
  printf 'Unable to read add-in Id from manifest: %s\n' "$MANIFEST_ABS" >&2
  exit 1
fi

SIDELOAD_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
TARGET_NAME="$ADDIN_ID.$(basename -- "$MANIFEST_ABS")"
TARGET_PATH="$SIDELOAD_DIR/$TARGET_NAME"

mkdir -p "$SIDELOAD_DIR"

for existing in "$SIDELOAD_DIR"/"$ADDIN_ID".*; do
  [ -e "$existing" ] || continue
  rm -f "$existing"
done

cp -f "$MANIFEST_ABS" "$TARGET_PATH"

printf 'Installed manifest: %s\n' "$TARGET_PATH"
printf 'Manifest source: %s\n' "$SOURCE_LOCATION"

if printf '%s' "$SOURCE_LOCATION" | grep -Eq 'https?://(localhost|127\.0\.0\.1)(:|/)'; then
  printf '%s\n' 'Warning: this manifest points to localhost. Use dist/manifest.xml for online deployment.'
else
  printf '%s\n' 'Warning: this manifest points to a remote URL. The add-in will only appear if that URL is reachable from PowerPoint.'
fi

printf '%s\n' 'Restart PowerPoint, open a presentation, then use Insert > My Add-ins next to the drop-down arrow to load the add-in.'
