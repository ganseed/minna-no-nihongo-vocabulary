#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")" && pwd)"
BUILD_ROOT="$PROJECT_ROOT/build/macos"
DIST_ROOT="$PROJECT_ROOT/dist/macos"
PACKAGE_ROOT="$DIST_ROOT/大家的日语词库-macOS"

mkdir -p "$PACKAGE_ROOT"
rm -rf "$PACKAGE_ROOT/大家的日语词库工具.app" "$PACKAGE_ROOT/data" "$PACKAGE_ROOT/template"
rm -f "$PACKAGE_ROOT/README.md"

python3 -m pip install --upgrade pip
python3 -m pip install -r "$PROJECT_ROOT/requirements.txt" pyinstaller

python3 -m PyInstaller --noconfirm --clean --windowed \
  --name "大家的日语词库工具" \
  --distpath "$PACKAGE_ROOT" \
  --workpath "$BUILD_ROOT" \
  --specpath "$BUILD_ROOT" \
  "$PROJECT_ROOT/desktop_app.py"

mkdir -p "$PACKAGE_ROOT/data" "$PACKAGE_ROOT/template" "$PACKAGE_ROOT/output"
cp "$PROJECT_ROOT/data/vocabulary.json" "$PACKAGE_ROOT/data/vocabulary.json"
cp "$PROJECT_ROOT/template/默写宏模板.xlsm" "$PACKAGE_ROOT/template/默写宏模板.xlsm"

echo "macOS 发布目录：$PACKAGE_ROOT"
