"""将词库导出为跨平台 PDF；每课从新页开始。"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def export_pdf(data_path: Path, output_path: Path) -> None:
    """从 JSON 导出包含词汇、固定表达与例句的 PDF。"""
    records = json.loads(data_path.read_text(encoding="utf-8"))
    grouped: dict[int, list[dict]] = defaultdict(list)
    for record in records:
        grouped[record["lesson"]].append(record)
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(str(output_path), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm)
    normal = ParagraphStyle("normal", fontName="STSong-Light", fontSize=10, leading=15, alignment=TA_CENTER)
    title = ParagraphStyle("title", parent=normal, fontSize=17, leading=22, spaceAfter=8 * mm)
    story = []
    for lesson, items in sorted(grouped.items()):
        if story:
            story.append(PageBreak())
        story.append(Paragraph(f"Lesson {lesson:02d}", title))
        regular = [item for item in items if item["type"] != "example"]
        examples = [item for item in items if item["type"] == "example"]
        table = Table([[Paragraph(item["japanese"].replace("\n", "<br/>"), normal), Paragraph(item["chinese"], normal)] for item in regular], colWidths=[110 * mm, 55 * mm])
        table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C3D0")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")])]))
        story.append(table)
        if examples:
            story.append(Spacer(1, 6 * mm))
            story.append(Paragraph("【例句】", title))
            for item in examples:
                story.append(Paragraph(item["japanese"].replace("\n", "<br/>") + "<br/>" + item["chinese"], normal))
                story.append(Spacer(1, 2 * mm))
    document.build(story)


if __name__ == "__main__":
    export_pdf(PROJECT_ROOT / "data" / "vocabulary.json", PROJECT_ROOT / "output" / "大家的日语Ⅰ词库.pdf")
