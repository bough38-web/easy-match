#!/usr/bin/env bash
set -euo pipefail
VERSION="4.7.0"
APP="dist/ExcelMatcherUniversal.app"
mkdir -p release
[[ -d "$APP" ]] || { echo "Build first"; exit 1; }
DMG="release/ExcelMatcherUniversal_mac_${VERSION}.dmg"
TMPDIR="$(mktemp -d)"
cp -R "$APP" "$TMPDIR/ExcelMatcherUniversal.app"
hdiutil create -volname "ExcelMatcherUniversal" -srcfolder "$TMPDIR" -ov -format UDZO "$DMG"
rm -rf "$TMPDIR"
echo "Created: $DMG"
