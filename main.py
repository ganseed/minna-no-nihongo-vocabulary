"""日语词库工程统一命令入口。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from scripts.excel_builder import build_workbooks
from scripts.export_word import export_word
from scripts.random_exam import DIRECTIONS, load_dataset, parse_count, parse_lessons


ROOT = Path(__file__).resolve().parent


def main() -> None:
    # 英文 Windows 的重定向输出可能默认使用 cp1252，中文路径同样需要 UTF-8。
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="大家的日语长期词库工程")
    parser.add_argument("command", choices=["excel", "word", "all"])
    parser.add_argument("--data", type=Path, default=ROOT / "data" / "vocabulary.json")
    parser.add_argument("--output", type=Path, default=ROOT / "output")
    parser.add_argument("--lessons", default="全部", help="课程，例如 3、1,3,5、1-5 或 全部")
    parser.add_argument("--count", default="20", help="正整数或 全部（只影响默写）")
    parser.add_argument("--direction", choices=list(DIRECTIONS), default="中译日")
    parser.add_argument("--seed", type=int, help="可选随机种子，仅用于复现试卷")
    args = parser.parse_args()
    try:
        records, metadata = load_dataset(args.data)
        lessons = parse_lessons(args.lessons, (r["lesson"] for r in records))
        if args.command in {"excel", "all"}:
            paths = build_workbooks(args.data, args.output, lessons=lessons,
                                    count=parse_count(args.count), direction=DIRECTIONS[args.direction], seed=args.seed)
            for path in paths:
                print(f"已生成：{path}")
        if args.command in {"word", "all"}:
            path = export_word(args.data, args.output / f"{metadata['title']}词库.docx", lessons=lessons)
            print(f"已生成：{path}")
    except (OSError, UnicodeError, ValueError) as error:
        parser.exit(1, f"生成失败：{error}\n")


if __name__ == "__main__":
    main()
