"""从统一 JSON 词库中抽取随机默写题。"""

from __future__ import annotations

import json
import random
from pathlib import Path
from typing import Iterable


def load_vocabulary(path: Path) -> list[dict]:
    """读取并检查统一词库。"""
    records = json.loads(path.read_text(encoding="utf-8"))
    required = {"id", "lesson", "order", "type", "japanese", "kana", "kanji", "chinese", "remark"}
    if any(required - record.keys() for record in records):
        raise ValueError("词库字段不完整。")
    return records


def select_questions(
    records: Iterable[dict],
    lessons: Iterable[int] | None = None,
    count: int = 20,
    direction: str = "jp_to_zh",
    seed: int | None = None,
) -> list[dict]:
    """按课次和方向随机抽题；例句默认不进入默写题。"""
    lesson_set = set(lessons or [])
    pool = [
        record
        for record in records
        if record["type"] in {"word", "expression"}
        and record["chinese"]
        and (not lesson_set or record["lesson"] in lesson_set)
    ]
    rng = random.Random(seed)
    chosen = rng.sample(pool, min(count, len(pool)))
    for number, record in enumerate(chosen, 1):
        record = record.copy()
        if direction == "jp_to_zh":
            record.update(number=number, question=record["japanese"], answer=record["chinese"])
        elif direction == "zh_to_jp":
            record.update(number=number, question=record["chinese"], answer=record["japanese"])
        else:
            raise ValueError("direction 必须是 jp_to_zh 或 zh_to_jp。")
        chosen[number - 1] = record
    return chosen
