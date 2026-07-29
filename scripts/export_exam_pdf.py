"""直接从 JSON 生成默写 PDF，避免打印 Excel 预留空行。"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle

from scripts.random_exam import load_vocabulary, select_questions


def export_exam_pdf(
    data_path: Path,
    output_path: Path,
    lessons: list[int] | None = None,
    count: int | None = 20,
    direction: str = "jp_to_zh",
    include_answers: bool = False,
    seed: int | None = None,
) -> None:
    """按实际抽题数量生成默写卷，可选在末尾附答案。"""
    records = load_vocabulary(data_path)
    available = [
        record
        for record in records
        if record["type"] in {"word", "expression"}
        and record.get("chinese")
        and (not lessons or record["lesson"] in set(lessons))
    ]
    requested = len(available) if count is None else count
    questions = select_questions(records, lessons, requested, direction, seed)
    if not questions:
        raise ValueError("所选课程没有可用于默写的单词或固定表达。")

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(output_path), pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=14 * mm, bottomMargin=14 * mm,
    )
    normal = ParagraphStyle(
        "exam-normal", fontName="STSong-Light", fontSize=10,
        leading=14, alignment=TA_CENTER,
    )
    title = ParagraphStyle(
        "exam-title", parent=normal, fontSize=16, leading=22, spaceAfter=5 * mm,
    )
    scope = "全部课程" if not lessons else "Lesson " + ", ".join(map(str, lessons))
    direction_name = "日译中" if direction == "jp_to_zh" else "中译日"
    story = [Paragraph(f"大家的日语Ⅰ 默写 - {direction_name}", title), Paragraph(f"范围：{scope}　题数：{len(questions)}", normal)]
    rows = [[Paragraph("序号", normal), Paragraph("题目", normal), Paragraph("作答", normal)]]
    for item in questions:
        question = str(item["question"]).replace("\n", "<br/>")
        rows.append([Paragraph(str(item["number"]), normal), Paragraph(question, normal), ""])
    table = Table(rows, colWidths=[16 * mm, 82 * mm, 82 * mm], repeatRows=1, rowHeights=[9 * mm] + [13 * mm] * len(questions))
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#375A7F")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4CE")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(table)

    if include_answers:
        story.extend([PageBreak(), Paragraph("参考答案", title)])
        answer_rows = [[Paragraph("序号", normal), Paragraph("答案", normal)]]
        for item in questions:
            answer = str(item["answer"]).replace("\n", "<br/>")
            answer_rows.append([Paragraph(str(item["number"]), normal), Paragraph(answer, normal)])
        answer_table = Table(answer_rows, colWidths=[20 * mm, 160 * mm], repeatRows=1)
        answer_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#375A7F")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4CE")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ]))
        story.append(answer_table)
    document.build(story)
