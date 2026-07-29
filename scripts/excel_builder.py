"""使用公开的 openpyxl 从统一 JSON 数据源生成 Excel。"""

from __future__ import annotations

import json
import random
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


# 打包为 EXE 后，以 EXE 所在目录为工程根目录，确保 JSON 始终可外部编辑。
PROJECT_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
DATA_FILE = PROJECT_ROOT / "data" / "vocabulary.json"
OUTPUT_DIR = PROJECT_ROOT / "output"

THIN = Side(style="thin", color="B8C4CE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="375A7F")
EXAMPLE_FILL = PatternFill("solid", fgColor="D88C4A")
HEADER_FONT = Font(name="Yu Gothic", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Yu Gothic", size=11, color="172B3A")
ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def load_records() -> list[dict]:
    """读取并校验词库的必要字段。"""
    records = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    required = {"id", "lesson", "order", "type", "japanese", "chinese"}
    for index, record in enumerate(records, 1):
        missing = required - record.keys()
        if missing:
            raise ValueError(f"第 {index} 条记录缺少字段：{', '.join(sorted(missing))}")
    return records


def style_range(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    """统一设置正文单元格样式。"""
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = ALIGN
            cell.border = BORDER


def style_header(ws, row: int, min_col: int, max_col: int, fill=HEADER_FILL) -> None:
    """设置标题行样式。"""
    for cells in ws.iter_rows(min_row=row, max_row=row, min_col=min_col, max_col=max_col):
        for cell in cells:
            cell.font = HEADER_FONT
            cell.fill = fill
            cell.alignment = ALIGN
            cell.border = BORDER


def prepare_workbook() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass
    return wb


def build_vocabulary(records: list[dict]) -> Path:
    """生成阅读词库和程序读取用 Vocabulary 工作表。"""
    wb = prepare_workbook()
    max_lesson = max(int(record["lesson"]) for record in records)
    for lesson in range(1, max_lesson + 1):
        ws = wb.create_sheet(f"Lesson{lesson:02d}")
        ws.append(["日语", "中文"])
        lesson_records = [record for record in records if int(record["lesson"]) == lesson]
        regular = [record for record in lesson_records if record["type"] != "example"]
        examples = [record for record in lesson_records if record["type"] == "example"]
        for record in regular:
            ws.append([record.get("japanese", ""), record.get("chinese", "")])
        if examples:
            example_row = ws.max_row + 1
            ws.append(["【例句】", ""])
            for record in examples:
                ws.append([record.get("japanese", ""), record.get("chinese", "")])
            style_header(ws, example_row, 1, 2, EXAMPLE_FILL)
        style_range(ws, 2, ws.max_row, 1, 2)
        style_header(ws, 1, 1, 2)
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 22
        for row in range(1, ws.max_row + 1):
            ws.row_dimensions[row].height = 38 if row > 1 else 28
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False

    ws = wb.create_sheet("Vocabulary")
    headers = ["id", "lesson", "order", "type", "japanese", "kana", "kanji", "chinese", "remark"]
    ws.append(headers)
    for record in records:
        ws.append([record.get(key, "") for key in headers])
    style_range(ws, 2, ws.max_row, 1, len(headers))
    style_header(ws, 1, 1, len(headers))
    widths = [16, 10, 10, 14, 34, 24, 24, 38, 18]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    path = OUTPUT_DIR / "大家的日语Ⅰ词库.xlsx"
    wb.save(path)
    return path


def build_exam(records: list[dict]) -> Path:
    """生成支持任意课程组合、方向和数量的默写工作簿。"""
    eligible = [r for r in records if r.get("type") in {"word", "expression"} and r.get("chinese")]
    random.shuffle(eligible)
    wb = prepare_workbook()

    settings = wb.create_sheet("设置")
    rows = [
        ["默写设置", "选择值"], ["方向", "日译中"], ["课程", "1"], ["随机数量", 20],
        ["刷新方式", "重新运行一键生成脚本"], ["随机种子", "每次生成自动刷新"],
        ["可选方向", "日译中 / 中译日"], ["课程填写示例", "单课：3；多课：1,3,5；全部：全部"],
        ["可选数量", "20 / 50 / 100 / 200 / 全部"],
        ["题目范围", "所选课程的全部单词和固定表达（不含例句）"],
    ]
    for row in rows:
        settings.append(row)
    style_range(settings, 2, len(rows), 1, 2)
    style_header(settings, 1, 1, 2)
    settings.column_dimensions["A"].width = 24
    settings.column_dimensions["B"].width = 48
    direction = DataValidation(type="list", formula1='"日译中,中译日"')
    amount = DataValidation(type="list", formula1='"20,50,100,200,全部"')
    settings.add_data_validation(direction)
    settings.add_data_validation(amount)
    direction.add(settings["B2"])
    amount.add(settings["B4"])
    settings.sheet_view.showGridLines = False

    source = wb.create_sheet("题库")
    source.append(["lesson", "japanese", "chinese", "type", "id"])
    for r in eligible:
        source.append([r["lesson"], r["japanese"], r["chinese"], r["type"], r["id"]])
    style_range(source, 2, source.max_row, 1, 5)
    style_header(source, 1, 1, 5)
    source.freeze_panes = "A2"
    source.sheet_view.showGridLines = False

    groups = wb.create_sheet("随机题组")
    groups.append(["lesson", "japanese", "chinese", "selected", "selected_order"])
    for r in eligible:
        groups.append([r["lesson"], r["japanese"], r["chinese"], "", ""])
    last_row = groups.max_row
    normalized = 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(\'设置\'!$B$3,"，",","),"、",",")," ","")'
    for row in range(2, last_row + 1):
        groups.cell(row, 4, f'=IF(\'设置\'!$B$3="全部",1,--ISNUMBER(SEARCH(","&A{row}&",",","&{normalized}&",")))')
        groups.cell(row, 5, f'=IF(D{row}=1,COUNTIF($D$2:D{row},1),"")')
    style_range(groups, 2, last_row, 1, 5)
    style_header(groups, 1, 1, 5)
    groups.sheet_state = "hidden"

    exam = wb.create_sheet("默写")
    exam.append(["序号", "题目", "作答", "答案"])
    limit = f'IF(\'设置\'!$B$4="全部",{len(eligible)},\'设置\'!$B$4)'
    for number in range(1, len(eligible) + 1):
        row = number + 1
        source_row = f'IFERROR(MATCH(A{row},\'随机题组\'!$E$2:$E${last_row},0),"")'
        question = f'=IF(OR(A{row}>{limit},{source_row}=""),"",INDEX(\'随机题组\'!$B$2:$C${last_row},{source_row},IF(\'设置\'!$B$2="日译中",1,2)))'
        answer = f'=IF(OR(A{row}>{limit},{source_row}=""),"",INDEX(\'随机题组\'!$B$2:$C${last_row},{source_row},IF(\'设置\'!$B$2="日译中",2,1)))'
        exam.append([number, question, "", answer])
    style_range(exam, 2, exam.max_row, 1, 4)
    style_header(exam, 1, 1, 4)
    for column, width in {"A": 9, "B": 40, "C": 30, "D": 34}.items():
        exam.column_dimensions[column].width = width
    for row in range(1, exam.max_row + 1):
        exam.row_dimensions[row].height = 36
    exam.freeze_panes = "A2"
    exam.sheet_view.showGridLines = False
    path = OUTPUT_DIR / "大家的日语Ⅰ默写.xlsx"
    wb.save(path)
    return path


def build_workbooks() -> None:
    """从 JSON 重新生成两个 Excel 文件。"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    records = load_records()
    build_vocabulary(records)
    build_exam(records)
    print(f"已生成 {len(records)} 条记录，输出目录：{OUTPUT_DIR}")


if __name__ == "__main__":
    build_workbooks()
