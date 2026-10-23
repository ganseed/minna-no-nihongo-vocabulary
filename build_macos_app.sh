#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")" && pwd)"
BUILD_ROOT="$PROJECT_ROOT/build/macos"
DIST_ROOT="$PROJECT_ROOT/dist/macos"
PACKAGE_ROOT="$DIST_ROOT/大家的日语词库-macOS"
APP_PATH="$PACKAGE_ROOT/大家的日语词库工具.app"
APP_RESOURCES="$APP_PATH/Contents/Resources"

mkdir -p "$PACKAGE_ROOT"
rm -rf "$APP_PATH" "$PACKAGE_ROOT/data" "$PACKAGE_ROOT/template"
rm -f "$PACKAGE_ROOT/README.md"

python3 -m pip install --upgrade pip
python3 -m pip install -r "$PROJECT_ROOT/requirements.txt" pyinstaller

python3 -m PyInstaller --noconfirm --clean --windowed \
  --name "大家的日语词库工具" \
  --add-data "$PROJECT_ROOT/data:data" \
  --add-data "$PROJECT_ROOT/template:template" \
  --distpath "$PACKAGE_ROOT" \
  --workpath "$BUILD_ROOT" \
  --specpath "$BUILD_ROOT" \
  "$PROJECT_ROOT/desktop_app.py"

# Do not rely only on PyInstaller's --add-data layout: place required resources
# explicitly in the standard macOS app bundle resource directory. This keeps
# the app self-contained when Gatekeeper/App Translocation relocates the .app.
mkdir -p "$APP_RESOURCES/data" "$APP_RESOURCES/template"
cp "$PROJECT_ROOT/data/vocabulary.json" "$APP_RESOURCES/data/vocabulary.json"
cp "$PROJECT_ROOT/template/默写宏模板.xlsm" "$APP_RESOURCES/template/默写宏模板.xlsm"

# Also ship editable copies next to the app for users who want to customize them.
mkdir -p "$PACKAGE_ROOT/data" "$PACKAGE_ROOT/template" "$PACKAGE_ROOT/output"
cp "$PROJECT_ROOT/data/vocabulary.json" "$PACKAGE_ROOT/data/vocabulary.json"
cp "$PROJECT_ROOT/template/默写宏模板.xlsm" "$PACKAGE_ROOT/template/默写宏模板.xlsm"

test -f "$APP_RESOURCES/data/vocabulary.json"
test -f "$APP_RESOURCES/template/默写宏模板.xlsm"

echo "macOS 发布目录：$PACKAGE_ROOT"
