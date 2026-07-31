"""日语词库工程统一命令入口。"""

from __future__ import annotations

import argparse
from pathlib import Path

from scripts.excel_builder import build_workbooks
from scripts.export_word import export_word


ROOT = Path(__file__).resolve().parent


def main() -> None:
    parser = argparse.ArgumentParser(description="大家的日语长期词库工程")
    parser.add_argument("command", choices=["excel", "word", "all"])
    parser.add_argument("--data", type=Path, default=ROOT / "data" / "vocabulary.json")
    parser.add_argument("--output", type=Path, default=ROOT / "output")
    args = parser.parse_args()
    if args.command in {"excel", "all"}:
        build_workbooks(args.data, args.output)
    if args.command in {"word", "all"}:
        export_word(args.data, args.output / "大家的日语Ⅰ词库.docx")


if __name__ == "__main__":
    main()
