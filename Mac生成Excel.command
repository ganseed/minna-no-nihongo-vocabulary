#!/bin/bash

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJECT_DIR" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 Python 3。请先从 https://www.python.org/downloads/macos/ 安装 Python 3。"
  read -r -p "按回车键关闭窗口……"
  exit 1
fi

PYTHONPATH="$PROJECT_DIR/vendor${PYTHONPATH:+:$PYTHONPATH}" python3 generate_excel.py --no-pause
STATUS=$?

if [ $STATUS -eq 0 ]; then
  echo ""
  echo "生成完成，请查看 output 文件夹。"
else
  echo ""
  echo "生成失败，请保留窗口中的错误信息。"
fi

read -r -p "按回车键关闭窗口……"
exit $STATUS
