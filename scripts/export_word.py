"""将统一词库导出为 Word 文档。"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def export_word(data_path: Path, output_path: Path) -> Path:
    """生成每课分节、例句置后的可编辑 Word 文档。"""
    records = json.loads(data_path.read_text(encoding="utf-8"))
    grouped: dict[int, list[dict]] = defaultdict(list)
    for record in records:
        grouped[record["lesson"]].append(record)
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "Yu Gothic"
    normal.font.size = Pt(10.5)
    for lesson, items in sorted(grouped.items()):
        if lesson > 1:
            document.add_page_break()
        heading = document.add_heading(f"Lesson {lesson:02d}", level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        regular = [item for item in items if item["type"] != "example"]
        examples = [item for item in items if item["type"] == "example"]
        table = document.add_table(rows=0, cols=2)
        table.style = "Table Grid"
        for item in regular:
            cells = table.add_row().cells
            cells[0].width, cells[1].width = Cm(12), Cm(6.6)
            cells[0].text, cells[1].text = item["japanese"], item["chinese"]
            for cell in cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if examples:
            document.add_heading("【例句】", level=2)
            for item in examples:
                paragraph = document.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.add_run(item["japanese"] + "\n" + item["chinese"])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


if __name__ == "__main__":
    export_word(PROJECT_ROOT / "data" / "vocabulary.json", PROJECT_ROOT / "output" / "大家的日语Ⅰ词库.docx")
