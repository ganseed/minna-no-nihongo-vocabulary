"""词库校验、课程解析与随机抽题的统一入口。"""
from __future__ import annotations
import json
import random
import re
from pathlib import Path
from typing import Iterable

DIRECTIONS = {"中译日": "zh_to_jp", "日译中": "jp_to_zh"}


def load_dataset(path: Path) -> tuple[list[dict], dict]:
    """兼容 UTF-8 BOM，报出具体出错记录，保留词库扩展字段。"""
    try:
        document = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise ValueError(f"JSON 第 {error.lineno} 行、第 {error.colno} 列格式有误。") from error
    metadata = document if isinstance(document, dict) else {}
    records = metadata.get("records") if metadata else document
    if not isinstance(records, list) or not records:
        raise ValueError("词库须为非空数组，或包含非空 records 数组的对象。")
    stem = Path(path).stem
    title = metadata.get("title", "大家的日语Ⅰ" if stem == "vocabulary" else stem)
    unit_label = metadata.get("unit_label", "课程")
    if not isinstance(title, str) or not title.strip() or re.search(r'[\\/:*?"<>|]', title) or len(title) > 80:
        raise ValueError("词库 title 应为1～80字的名称，不能含文件路径或文件名保留符号。")
    if not isinstance(unit_label, str) or not unit_label.strip() or len(unit_label) > 10 or re.search(r'[\[\]\\/*?:]', unit_label):
        raise ValueError("unit_label 应为1～10字，例如“课程”“主题”或“分组”。")
    metadata = {**metadata, "title": title.strip(), "unit_label": unit_label.strip()}
    required = {"id", "japanese", "chinese"}
    seen = set()
    for index, record in enumerate(records, 1):
        prefix = f"第 {index} 条记录"
        if not isinstance(record, dict):
            raise ValueError(f"{prefix}必须是对象。")
        missing = required - record.keys()
        if missing:
            raise ValueError(f"{prefix}缺少字段：{', '.join(sorted(missing))}")
        record.setdefault("lesson", 1)
        record.setdefault("order", index)
        record.setdefault("type", "word")
        for key in ("id", "type", "japanese", "chinese"):
            if not isinstance(record[key], str):
                raise ValueError(f"{prefix}的 {key} 必须是文字。")
        if not record["id"].strip() or not record["japanese"].strip():
            raise ValueError(f"{prefix}的 id 和 japanese 不能为空。")
        if record["id"] in seen:
            raise ValueError(f"{prefix}的 id 重复：{record['id']}")
        seen.add(record["id"])
        for key in ("lesson", "order"):
            if type(record[key]) is not int or record[key] < 1:
                raise ValueError(f"{prefix}的 {key} 必须是正整数。")
    return records, metadata


def load_vocabulary(path: Path) -> list[dict]:
    return load_dataset(path)[0]


def parse_lessons(text: str, available: Iterable[int]) -> list[int] | None:
    """支持 3、1,3,5、1-5 及中文分隔符；拒绝不存在的课次。"""
    available = set(available)
    text = text.strip()
    if text in {"", "全部"}:
        return None
    normalized = re.sub(r"[，、;；]", ",", text)
    selected = set()
    for token in normalized.split(","):
        match = re.fullmatch(r"\s*(\d+)\s*(?:[-～~—]\s*(\d+)\s*)?", token)
        if not match:
            raise ValueError("课程请填写 3、1,3,5、1-5 或“全部”。")
        first = int(match[1])
        last = int(match[2] or match[1])
        if first > last:
            raise ValueError("课程范围应从小到大，例如 1-5。")
        if first < 1 or last > max(available, default=0):
            raise ValueError("所选课程超出了词库的课次范围。")
        selected.update(range(first, last + 1))
    missing = selected - available
    if missing:
        raise ValueError(f"词库中没有这些课程：{', '.join(map(str, sorted(missing)))}")
    return sorted(selected)


def parse_count(value: str | int | None) -> int | None:
    if value is None or value == "全部":
        return None
    if type(value) is int and value > 0:
        return value
    if isinstance(value, str) and re.fullmatch(r"[1-9]\d*", value.strip()):
        return int(value.strip())
    raise ValueError("题数请填写正整数或“全部”。")


def filter_records(records: Iterable[dict], lessons: Iterable[int] | None = None) -> list[dict]:
    selected = None if lessons is None else set(lessons)
    if selected == set():
        raise ValueError("请至少选择一课。")
    return [record for record in records if selected is None or record["lesson"] in selected]


def select_questions(records: Iterable[dict], lessons: Iterable[int] | None = None,
                     count: int | None = 20, direction: str = "zh_to_jp",
                     seed: int | None = None) -> list[dict]:
    """先筛选再无放回抽题；全部和题数不足时均不生成占位空题。"""
    if direction not in DIRECTIONS.values():
        raise ValueError("direction 必须是 jp_to_zh 或 zh_to_jp。")
    count = parse_count(count)
    pool = [r for r in filter_records(records, lessons) if r["chinese"].strip()]
    if not pool:
        raise ValueError("所选课程没有带中文释义的可用题目。")
    chosen = random.Random(seed).sample(pool, len(pool) if count is None else min(count, len(pool)))
    result = []
    for number, original in enumerate(chosen, 1):
        record = original.copy()
        question, answer = (record["japanese"], record["chinese"])
        if direction == "zh_to_jp":
            question, answer = answer, question
        record.update(number=number, question=question, answer=answer)
        result.append(record)
    return result
