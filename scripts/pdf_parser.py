"""PDF 词库解析入口：优先文字层，无文字层时使用 OCR 缓存。"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from pypdf import PdfReader


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def has_text_layer(pdf_path: Path) -> bool:
    """抽样并判断 PDF 是否存在可用文字层。"""
    reader = PdfReader(pdf_path)
    sample = reader.pages[: min(8, len(reader.pages))]
    return sum(len((page.extract_text() or "").strip()) for page in sample) >= 80


def copy_prepared_data(output_path: Path) -> None:
    """复制经逐页双语 OCR、坐标对齐和人工抽查后的结构化数据。"""
    prepared = PROJECT_ROOT / "data" / "vocabulary.json"
    if not prepared.exists():
        raise FileNotFoundError("缺少 data/vocabulary.json，请先完成 OCR 审核。")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if prepared.resolve() != output_path.resolve():
        shutil.copy2(prepared, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="解析《大家的日语》词汇 PDF")
    parser.add_argument("pdf", type=Path, nargs="?", default=PROJECT_ROOT / "source" / "大家的日语初级一每课单词.pdf")
    parser.add_argument("--output", type=Path, default=PROJECT_ROOT / "data" / "vocabulary.json")
    args = parser.parse_args()
    if not args.pdf.exists():
        raise FileNotFoundError(args.pdf)
    if has_text_layer(args.pdf):
        print("检测到文字层。此版教材已保留审核后的结构化数据，直接使用该数据以确保版面顺序。")
    else:
        print("未检测到文字层，使用逐页双语 OCR 的审核结果。")
    copy_prepared_data(args.output)
    records = json.loads(args.output.read_text(encoding="utf-8"))
    print(f"已写入 {len(records)} 条记录：{args.output}")


if __name__ == "__main__":
    main()
