#!/bin/bash
# Build "Assessment PDF to Excel.app" and a drag-to-install .dmg.
# Usage (from anywhere):  bash installer/build_mac.sh
set -euo pipefail
cd "$(dirname "$0")/.."

APP_NAME="Assessment PDF to Excel"
PYTHON="${PYTHON:-.venv/bin/python}"

"$PYTHON" -m pip install --quiet -r requirements.txt -r installer/requirements-build.txt 2>/dev/null \
  || uv pip install --python "$PYTHON" --quiet -r requirements.txt -r installer/requirements-build.txt

rm -rf build dist
"$PYTHON" -m PyInstaller installer/converter.spec --noconfirm --clean

# Ad-hoc signature so macOS accepts the bundle as one unit.
codesign --force --deep --sign - "dist/$APP_NAME.app"

STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
cp -R "dist/$APP_NAME.app" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
hdiutil create -volname "$APP_NAME" -srcfolder "$STAGE" -ov -format UDZO \
  "dist/Assessment-PDF-to-Excel-mac.dmg" >/dev/null

echo "Built: dist/Assessment-PDF-to-Excel-mac.dmg"
