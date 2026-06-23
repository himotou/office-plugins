#!/bin/sh

set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
REPO_DIR=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
MANIFEST_PATH="$REPO_DIR/manifest.xml"

if [ ! -f "$MANIFEST_PATH" ]; then
  printf 'Development manifest not found: %s\n' "$MANIFEST_PATH" >&2
  exit 1
fi

printf 'Installing development manifest: %s\n' "$MANIFEST_PATH"
sh "$SCRIPT_DIR/install-office-addin-mac.sh" "$MANIFEST_PATH"
