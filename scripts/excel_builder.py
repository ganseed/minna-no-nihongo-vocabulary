"""使用公开的 openpyxl 从统一 JSON 数据源生成 Excel。"""

from __future__ import annotations

import json
import random
import sys
import tempfile
import zipfile
from pathlib import Path

import xlsxwriter
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


def project_root() -> Path:
    """定位应用旁边的外部 data 目录。"""
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve().parents[1]
    executable = Path(sys.executable).resolve()
    # macOS 应用位于发布目录/App.app/Contents/MacOS/，数据放在 App.app 旁边。
    if sys.platform == "darwin" and executable.parent.name == "MacOS":
        return executable.parents[3]
    return executable.parent


PROJECT_ROOT = project_root()
DATA_FILE = PROJECT_ROOT / "data" / "vocabulary.json"
OUTPUT_DIR = PROJECT_ROOT / "output"
MACRO_TEMPLATE = PROJECT_ROOT / "template" / "默写宏模板.xlsm"

THIN = Side(style="thin", color="B8C4CE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="375A7F")
EXAMPLE_FILL = PatternFill("solid", fgColor="D88C4A")
HEADER_FONT = Font(name="Yu Gothic", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Yu Gothic", size=11, color="172B3A")
ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def load_records(data_path: Path | None = None) -> list[dict]:
    """读取并校验词库的必要字段。"""
    source = Path(data_path) if data_path else DATA_FILE
    records = json.loads(source.read_text(encoding="utf-8"))
    if not isinstance(records, list) or not records:
        raise ValueError("词库 JSON 必须是非空数组。")
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


def build_vocabulary(records: list[dict], output_dir: Path | None = None) -> Path:
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
    destination = Path(output_dir) if output_dir else OUTPUT_DIR
    destination.mkdir(parents=True, exist_ok=True)
    path = destination / "大家的日语Ⅰ词库.xlsx"
    wb.save(path)
    return path


def build_exam(records: list[dict], output_dir: Path | None = None) -> Path:
    """生成支持任意课程组合、方向和数量的默写工作簿。"""
    eligible = [r for r in records if r.get("chinese")]
    random.shuffle(eligible)
    wb = prepare_workbook()

    settings = wb.create_sheet("设置")
    rows = [
        ["默写设置", "选择值"], ["方向", "中译日"], ["课程", "全部"], ["随机数量", 20],
        ["刷新方式", "重新运行一键生成脚本"], ["随机种子", "每次生成自动刷新"],
        ["可选方向", "日译中 / 中译日"], ["课程填写示例", "单课：3；多课：1,3,5；全部：全部"],
        ["可选数量", "20 / 50 / 100 / 200 / 全部"],
        ["题目范围", "所选课程的全部记录"],
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
    destination = Path(output_dir) if output_dir else OUTPUT_DIR
    destination.mkdir(parents=True, exist_ok=True)
    path = destination / "大家的日语Ⅰ默写.xlsx"
    wb.save(path)
    return path


def build_macro_exam(records: list[dict], output_dir: Path | None = None) -> Path:
    """使用内置 VBA 创建跨平台可生成的宏版默写工作簿。"""
    if not MACRO_TEMPLATE.is_file():
        raise FileNotFoundError(f"缺少宏模板：{MACRO_TEMPLATE}")

    destination = Path(output_dir) if output_dir else OUTPUT_DIR
    destination.mkdir(parents=True, exist_ok=True)
    path = destination / "大家的日语Ⅰ默写-宏版.xlsm"
    eligible = [record for record in records if record.get("chinese")]
    random.shuffle(eligible)

    with zipfile.ZipFile(MACRO_TEMPLATE) as archive:
        vba_project = archive.read("xl/vbaProject.bin")

    with tempfile.TemporaryDirectory() as temp_dir:
        vba_path = Path(temp_dir) / "vbaProject.bin"
        vba_path.write_bytes(vba_project)
        workbook = xlsxwriter.Workbook(path)
        workbook.add_vba_project(vba_path)
        header = workbook.add_format({
            "font_name": "Yu Gothic", "font_size": 11, "bold": True,
            "font_color": "#FFFFFF", "bg_color": "#375A7F",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        body = workbook.add_format({
            "font_name": "Yu Gothic", "font_size": 11, "font_color": "#172B3A",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })

        settings = workbook.add_worksheet("设置")
        settings.hide_gridlines(2)
        settings.set_column("A:A", 24)
        settings.set_column("B:B", 48)
        settings_rows = [
            ["默写设置", "选择值"], ["方向", "中译日"], ["课程", "全部"], ["随机数量", 20],
            ["使用方式", "设置完成后点击右侧“随机生成”"],
            ["随机规则", "每次点击都会重新打乱"],
            ["可选方向", "日译中 / 中译日"],
            ["课程填写示例", "单课：3；多课：1,3,5；全部：全部"],
            ["可选数量", "20 / 50 / 100 / 200 / 全部"],
            ["题目范围", "所选课程的全部记录"],
        ]
        for row, values in enumerate(settings_rows):
            settings.write_row(row, 0, values, header if row == 0 else body)
        settings.data_validation("B2", {"validate": "list", "source": ["日译中", "中译日"]})
        settings.data_validation("B4", {"validate": "list", "source": [20, 50, 100, 200, "全部"]})
        settings.insert_button("D2", {
            "macro": "RegenerateExam", "caption": "随机生成",
            "width": 150, "height": 42,
        })

        source = workbook.add_worksheet("题库")
        source.hide_gridlines(2)
        source.freeze_panes(1, 0)
        source.set_column("A:A", 10)
        source.set_column("B:C", 36)
        source.set_column("D:E", 16)
        source.write_row(0, 0, ["lesson", "japanese", "chinese", "type", "id"], header)
        for row, record in enumerate(eligible, 1):
            source.write_row(row, 0, [
                record["lesson"], record["japanese"], record["chinese"],
                record["type"], record["id"],
            ], body)

        exam = workbook.add_worksheet("默写")
        exam.hide_gridlines(2)
        exam.freeze_panes(1, 0)
        exam.set_column("A:A", 9)
        exam.set_column("B:B", 40)
        exam.set_column("C:C", 30)
        exam.set_column("D:D", 34)
        exam.write_row(0, 0, ["序号", "题目", "作答", "答案"], header)
        # 预先为全部可用题目行设置统一边框，宏写入更多题目后样式仍然完整。
        for row in range(1, len(eligible) + 1):
            exam.set_row(row, 36)
            exam.write_blank(row, 0, None, body)
            exam.write_blank(row, 1, None, body)
            exam.write_blank(row, 2, None, body)
            exam.write_blank(row, 3, None, body)
        for row, record in enumerate(eligible[:20], 1):
            exam.write_row(row, 0, [row, record["chinese"], "", record["japanese"]], body)
        workbook.close()
    return path


def build_workbooks(
    data_path: Path | None = None,
    output_dir: Path | None = None,
) -> tuple[Path, Path, Path]:
    """从 JSON 生成词库、普通默写和宏版默写三个 Excel 文件。"""
    destination = Path(output_dir) if output_dir else OUTPUT_DIR
    destination.mkdir(parents=True, exist_ok=True)
    records = load_records(data_path)
    vocabulary_path = build_vocabulary(records, destination)
    exam_path = build_exam(records, destination)
    macro_exam_path = build_macro_exam(records, destination)
    print(f"已生成 {len(records)} 条记录，输出目录：{destination}")
    return vocabulary_path, exam_path, macro_exam_path


if __name__ == "__main__":
    build_workbooks()
