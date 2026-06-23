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

if [ ! -x "$REPO_DIR/node_modules/.bin/office-addin-dev-settings" ]; then
  printf '%s\n' 'office-addin-dev-settings is not installed. Run npm install first.' >&2
  exit 1
fi

printf '%s\n' 'Clearing stale PowerPoint sideload cache...'
sh "$SCRIPT_DIR/remove-office-addin-mac.sh" "$MANIFEST_ABS" >/dev/null 2>&1 || true

printf 'Registering add-in and launching PowerPoint with sideload document: %s\n' "$MANIFEST_ABS"
"$REPO_DIR/node_modules/.bin/office-addin-dev-settings" sideload "$MANIFEST_ABS" desktop --app powerpoint
