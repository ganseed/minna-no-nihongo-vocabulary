"""从统一词库生成阅读词库、固定试卷和可重新随机出题的宏版。"""
from __future__ import annotations
import math
import os
import secrets
import sys
import tempfile
import zipfile
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path

import xlsxwriter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from scripts.random_exam import DIRECTIONS, filter_records, load_vocabulary, select_questions


def project_root() -> Path:
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve().parents[1]
    executable = Path(sys.executable).resolve()
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
    return load_vocabulary(Path(data_path) if data_path else DATA_FILE)


@contextmanager
def atomic_output(path: Path):
    """写完临时文件再替换，生成失败不破坏已有文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=".generating-", suffix=path.suffix, dir=path.parent)
    os.close(handle)
    temporary = Path(temporary)
    try:
        yield temporary
        os.replace(temporary, path)
    except PermissionError as error:
        raise PermissionError(f"无法保存 {path.name}，请关闭正在打开的同名文件，或选择可写的输出文件夹。") from error
    finally:
        temporary.unlink(missing_ok=True)


def style_range(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.font, cell.alignment, cell.border = BODY_FONT, ALIGN, BORDER
            # 词库内容按原文显示，不能把以等号开头的词条解释为公式。
            if isinstance(cell.value, str):
                cell.data_type = "s"


def style_header(ws, row, min_col, max_col, fill=HEADER_FILL):
    for cells in ws.iter_rows(min_row=row, max_row=row, min_col=min_col, max_col=max_col):
        for cell in cells:
            cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, fill, ALIGN, BORDER


def prepare_workbook():
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def row_height(*values: str, minimum=36) -> float:
    """中日文按宽字符估算行高，避免较长释义被截断。"""
    lines = max(sum(max(1, math.ceil(sum(1 if ord(c) > 127 else .55 for c in line) / 27))
                    for line in str(value).split("\n")) for value in values)
    return max(minimum, lines * 15 + 6)


def print_layout(ws, last_row, last_col):
    ws.print_area = f"A1:{last_col}{last_row}"
    ws.print_title_rows = "1:1"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"


def build_vocabulary(records, output_dir=None):
    grouped = defaultdict(list)
    for record in sorted(records, key=lambda r: (r["lesson"], r["order"])):
        grouped[record["lesson"]].append(record)
    if not grouped:
        raise ValueError("所选课程没有词条。")
    wb = prepare_workbook()
    for lesson, items in sorted(grouped.items()):
        ws = wb.create_sheet(f"Lesson{lesson:02d}")
        ws.append(["日语", "中文"])
        examples = [r for r in items if r["type"] == "example"]
        for record in items:
            if record["type"] != "example":
                ws.append([record["japanese"], record["chinese"]])
        example_row = None
        if examples:
            example_row = ws.max_row + 1
            ws.append(["【例句】", ""])
            for record in examples:
                ws.append([record["japanese"], record["chinese"]])
        style_range(ws, 2, ws.max_row, 1, 2)
        style_header(ws, 1, 1, 2)
        if example_row:
            style_header(ws, example_row, 1, 2, EXAMPLE_FILL)
        ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 40, 40
        for row in ws:
            ws.row_dimensions[row[0].row].height = row_height(*(cell.value or "" for cell in row))
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        print_layout(ws, ws.max_row, "B")
    ws = wb.create_sheet("Vocabulary")
    headers = ["id", "lesson", "order", "type", "japanese", "kana", "kanji", "chinese", "remark",
               "part_of_speech", "pitch_accent", "dictionary_form"]
    ws.append(headers)
    for record in records:
        ws.append([record.get(key, "") for key in headers])
    style_range(ws, 2, ws.max_row, 1, len(headers))
    style_header(ws, 1, 1, len(headers))
    for index, width in enumerate([16, 10, 10, 14, 40, 24, 24, 40, 24, 16, 16, 24], 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    path = Path(output_dir or OUTPUT_DIR) / "大家的日语Ⅰ词库.xlsx"
    with atomic_output(path) as temporary:
        wb.save(temporary)
    return path


def build_exam(records, output_dir=None, *, lessons=None, count=20, direction="zh_to_jp", seed=None):
    """普通版是本次试卷快照，严格按实际题数建表，无空题占位。"""
    questions = select_questions(records, lessons, count, direction, seed)
    wb = prepare_workbook()
    for name, headers in [("默写", ["序号", "题目", "作答"]), ("答案", ["序号", "题目", "答案"])]:
        ws = wb.create_sheet(name)
        ws.append(headers)
        for q in questions:
            ws.append([q["number"], q["question"], "" if name == "默写" else q["answer"]])
            ws.row_dimensions[ws.max_row].height = row_height(q["question"], q["answer"])
        style_range(ws, 2, ws.max_row, 1, 3)
        style_header(ws, 1, 1, 3)
        for col, width in {"A": 9, "B": 48, "C": 38}.items():
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "B2"
        ws.sheet_view.showGridLines = False
        print_layout(ws, ws.max_row, "C")
    settings = wb.create_sheet("本次设置")
    label = next(k for k, v in DIRECTIONS.items() if v == direction)
    for values in [["本次试卷", "内容"], ["方向", label],
                   ["课程", "全部" if lessons is None else ",".join(map(str, lessons))],
                   ["实际题数", len(questions)],
                   ["更换课程或重新出题", "回到桌面工具重新生成；也可使用宏版的“设置”工作表。"],
                   ["打印提示", "打印“默写”工作表；答案在独立的“答案”工作表。"]]:
        settings.append(values)
    style_range(settings, 2, settings.max_row, 1, 2)
    style_header(settings, 1, 1, 2)
    settings.column_dimensions["A"].width, settings.column_dimensions["B"].width = 25, 65
    for row in range(2, settings.max_row + 1):
        settings.row_dimensions[row].height = 36
    settings.sheet_view.showGridLines = False
    path = Path(output_dir or OUTPUT_DIR) / "大家的日语Ⅰ默写.xlsx"
    with atomic_output(path) as temporary:
        wb.save(temporary)
    return path


def build_macro_exam(records, output_dir=None, *, lessons=None, count=20, direction="zh_to_jp", seed=None):
    """保持现有 VBA；用条件边框及动态打印范围适配多题和少题。"""
    if not MACRO_TEMPLATE.is_file():
        raise FileNotFoundError(f"缺少宏模板：{MACRO_TEMPLATE}")
    questions = select_questions(records, lessons, count, direction, seed)
    eligible = [r for r in records if r["chinese"].strip()]
    path = Path(output_dir or OUTPUT_DIR) / "大家的日语Ⅰ默写-宏版.xlsm"
    with zipfile.ZipFile(MACRO_TEMPLATE) as archive:
        vba_project = archive.read("xl/vbaProject.bin")
    with atomic_output(path) as temporary, tempfile.TemporaryDirectory() as temp_dir:
        vba_path = Path(temp_dir) / "vbaProject.bin"
        vba_path.write_bytes(vba_project)
        with xlsxwriter.Workbook(temporary, {"strings_to_formulas": False, "strings_to_urls": False}) as wb:
            wb.add_vba_project(vba_path)
            basic = {"font_name": "Yu Gothic", "font_size": 11, "font_color": "#172B3A",
                     "align": "center", "valign": "vcenter", "text_wrap": True}
            body = wb.add_format(basic)
            header = wb.add_format({**basic, "bold": True, "font_color": "#FFFFFF",
                                    "bg_color": "#375A7F", "border": 1})
            borders = wb.add_format({"border": 1, "border_color": "#B8C4CE"})
            settings = wb.add_worksheet("设置")
            source = wb.add_worksheet("题库")
            exam = wb.add_worksheet("默写")
            answers = wb.add_worksheet("答案")
            for sheet in (settings, source, exam, answers):
                sheet.hide_gridlines(2)
            settings.set_column("A:A", 24, body)
            settings.set_column("B:B", 60, body)
            settings.set_default_row(32)
            label = next(k for k, v in DIRECTIONS.items() if v == direction)
            settings_rows = [["默写设置", "选择值"], ["方向", label],
                ["课程", "全部" if lessons is None else ",".join(map(str, lessons))],
                ["随机数量", "全部" if count is None else count],
                ["使用方式", "修改设置后点击右侧“随机生成”；查看题目请切换到“默写”。"],
                ["打印提示", "默写与答案分别打印，范围随实际题数变化。"],
                ["课程填写示例", "单课：3；多课：1,3,5；全部：全部（此处不支持范围写法）"],
                ["实际题数", ""], ["题库范围", f"保留全部 {len(eligible)} 条可用题目，可继续更换课程。"]]
            for row, values in enumerate(settings_rows):
                settings.write_row(row, 0, values, header if row == 0 else body)
            settings.write_formula("B8", "=COUNTA('默写'!A:A)-1", body, len(questions))
            settings.data_validation("B2", {"validate": "list", "source": list(DIRECTIONS), "error_type": "stop", "error_message": "请选择中译日或日译中。", "ignore_blank": False})
            settings.data_validation("B4", {"validate": "list", "source": [20, 50, 100, 200, "全部"], "error_type": "stop", "error_message": "请选择列表中的题数。", "ignore_blank": False})
            settings.insert_button("D2", {"macro": "RegenerateExam", "caption": "随机生成", "width": 150, "height": 42})
            source.write_row(0, 0, ["lesson", "japanese", "chinese", "type", "id"], header)
            for row, record in enumerate(eligible, 1):
                source.write_row(row, 0, [record[k] for k in ["lesson", "japanese", "chinese", "type", "id"]], body)
            source.hide()
            for sheet in (exam, answers):
                sheet.set_column("A:A", 9, body)
                sheet.set_column("B:B", 48, body)
                sheet.set_column("C:C", 38, body)
                sheet.set_default_row(36)
                sheet.freeze_panes(1, 1)
                sheet.repeat_rows(0)
                sheet.set_paper(9)
                sheet.fit_to_pages(1, 0)
                sheet.set_footer("&C第 &P 页 / 共 &N 页")
                # 条件格式覆盖整个可用题数，但不创建占位单元格。
                # VBA 清除旧题后，旧边框随 A 列变空自动消失；增题时自动补齐。
                sheet.conditional_format(1, 0, len(eligible), 2,
                    {"type": "formula", "criteria": '$A2<>""', "format": borders})
            exam.set_column("D:D", 38, body, {"hidden": True})
            exam.write_row(0, 0, ["序号", "题目", "作答", "答案（见答案页）"], header)
            answers.write_row(0, 0, ["序号", "题目", "答案"], header)
            for row, q in enumerate(questions, 1):
                exam.write_row(row, 0, [row, q["question"], "", q["answer"]], body)
            # 预留答案公式不带固定边框；未出题的行显示空白，且不进入打印区域。
            for row in range(1, len(eligible) + 1):
                q = questions[row - 1] if row <= len(questions) else None
                for col, source_col, cached in [(0, "A", row if q else ""),
                    (1, "B", q["question"] if q else ""), (2, "D", q["answer"] if q else "")]:
                    answers.write_formula(row, col, f'=IF(\'默写\'!A{row+1}="","",\'默写\'!{source_col}{row+1})', body, cached)
            # Excel 原生动态打印区域：题数由 VBA 写入的 A 列决定，不受历史最多题数影响。
            for name in ("默写", "答案"):
                wb.define_name(f"'{name}'!_xlnm.Print_Area",
                    f"='{name}'!$A$1:INDEX('{name}'!$C:$C,MAX(2,COUNTA('默写'!$A:$A)))")
            exam.activate()
    return path


def build_workbooks(data_path=None, output_dir=None, *, lessons=None, count=20, direction="zh_to_jp", seed=None):
    records = load_records(data_path)
    selected = filter_records(records, lessons)
    # 一次生成中的普通版和宏版初始试卷保持一致；下次点击仍重新随机。
    seed = secrets.randbits(64) if seed is None else seed
    # 出题设置错误时，在写任何文件前立即提示。
    select_questions(records, lessons, count, direction, seed)
    destination = Path(output_dir or OUTPUT_DIR)
    destination.mkdir(parents=True, exist_ok=True)
    # 先完成全部生成，再替换正式输出，避免中途生成失败留下半套新文件。
    with tempfile.TemporaryDirectory(prefix=".generation-", dir=destination) as stage:
        stage = Path(stage)
        paths = (build_vocabulary(selected, stage),
                 build_exam(records, stage, lessons=lessons, count=count, direction=direction, seed=seed),
                 build_macro_exam(records, stage, lessons=lessons, count=count, direction=direction, seed=seed))
        result = []
        for source in paths:
            target = destination / source.name
            with atomic_output(target) as temporary:
                temporary.write_bytes(source.read_bytes())
            result.append(target)
    return tuple(result)


if __name__ == "__main__":
    build_workbooks()
