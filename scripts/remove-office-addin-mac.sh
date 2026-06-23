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

if [ -z "$ADDIN_ID" ]; then
  printf 'Unable to read add-in Id from manifest: %s\n' "$MANIFEST_ABS" >&2
  exit 1
fi

SIDELOAD_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
HOST_CACHE_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches"
HOST_WEF_CACHE_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
OSF_WEBHOST_DIR="$HOME/Library/Containers/com.Microsoft.OsfWebHost/Data"
SERVICE_V2_CACHE_DIR="$HOME/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2"
SERVICE_V2_LIBRARY_CACHE_DIR="$HOME/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2"
REMOVED=0

clear_dir_contents() {
  target_dir=$1

  if [ ! -d "$target_dir" ]; then
    return
  fi

  find "$target_dir" -mindepth 1 -maxdepth 1 -exec rm -rf {} +
  printf 'Cleared cache directory: %s\n' "$target_dir"
  REMOVED=1
}

clear_dir_contents "$SIDELOAD_DIR"

if [ -d "$OSF_WEBHOST_DIR" ]; then
  clear_dir_contents "$OSF_WEBHOST_DIR"
else
  clear_dir_contents "$HOST_CACHE_DIR"
  clear_dir_contents "$HOST_WEF_CACHE_DIR"
  clear_dir_contents "$SERVICE_V2_CACHE_DIR"
  clear_dir_contents "$SERVICE_V2_LIBRARY_CACHE_DIR"
fi

if [ "$REMOVED" -eq 0 ]; then
  printf 'No local Office cache found for add-in %s\n' "$ADDIN_ID"
fi

printf '%s\n' 'Restart PowerPoint, open a presentation, then use Insert > My Add-ins next to the drop-down arrow.'
