"""日语词库工程统一命令入口。"""

from __future__ import annotations

import argparse
from pathlib import Path

from scripts.excel_builder import build_workbooks
from scripts.export_pdf import export_pdf
from scripts.export_exam_pdf import export_exam_pdf
from scripts.export_word import export_word


ROOT = Path(__file__).resolve().parent


def main() -> None:
    parser = argparse.ArgumentParser(description="大家的日语长期词库工程")
    parser.add_argument("command", choices=["excel", "pdf", "exam-pdf", "word", "all"])
    parser.add_argument("--lessons", help="默写课次，如 1,3,5；省略表示全部课程")
    parser.add_argument("--count", default="20", help="默写数量：20、50、100、200 或 all")
    parser.add_argument("--direction", choices=["jp_to_zh", "zh_to_jp"], default="jp_to_zh")
    parser.add_argument("--answers", action="store_true", help="默写 PDF 末尾附参考答案")
    parser.add_argument("--seed", type=int, help="固定随机种子，便于复现同一套题")
    args = parser.parse_args()
    data = ROOT / "data" / "vocabulary.json"
    if args.command in {"excel", "all"}:
        build_workbooks()
    if args.command in {"pdf", "all"}:
        export_pdf(data, ROOT / "output" / "大家的日语Ⅰ词库.pdf")
    if args.command == "exam-pdf":
        lessons = [int(value.strip()) for value in args.lessons.split(",")] if args.lessons else None
        count = None if args.count.lower() == "all" else int(args.count)
        export_exam_pdf(
            data,
            ROOT / "output" / "大家的日语Ⅰ默写.pdf",
            lessons=lessons,
            count=count,
            direction=args.direction,
            include_answers=args.answers,
            seed=args.seed,
        )
    if args.command in {"word", "all"}:
        export_word(data, ROOT / "output" / "大家的日语Ⅰ词库.docx")


if __name__ == "__main__":
    main()
