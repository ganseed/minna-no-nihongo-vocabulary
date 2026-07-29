#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")" && pwd)"
BUILD_ROOT="$PROJECT_ROOT/build"
RELEASE_ROOT="$PROJECT_ROOT/dist/macos"

python3 -m pip install -r "$PROJECT_ROOT/requirements.txt" pyinstaller
python3 -m PyInstaller --noconfirm --clean --windowed \
  --name "大家的日语词库生成器" \
  --distpath "$RELEASE_ROOT" \
  --workpath "$BUILD_ROOT/macos" \
  --specpath "$BUILD_ROOT" \
  "$PROJECT_ROOT/desktop_app.py"

mkdir -p "$RELEASE_ROOT/data" "$RELEASE_ROOT/output"
cp "$PROJECT_ROOT/data/vocabulary.json" "$RELEASE_ROOT/data/vocabulary.json"
cp "$PROJECT_ROOT/应用使用说明.txt" "$RELEASE_ROOT/使用说明.txt"
echo "macOS 应用已生成：$RELEASE_ROOT"
