# -*- coding: utf-8 -*-
"""
文字檔解析器：讀取 UPDB 批次加人輸入檔，輸出結構化紀錄。
格式見 docs/INPUT_FORMAT.md。
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import List

logger = logging.getLogger(__name__)


@dataclass
class Record:
    """單筆解析結果。"""
    project: str
    group: str
    add_svn: bool
    employee_id: str
    line_number: int = 0  # 方便錯誤報告

    def __repr__(self) -> str:
        return (
            f"Record(project={self.project!r}, group={self.group!r}, "
            f"add_svn={self.add_svn}, employee_id={self.employee_id!r})"
        )


def parse_line(line: str, line_number: int) -> Record | None:
    """
    解析單行。以 Tab 分第一欄/第二欄；第一欄以空白分 token。
    第 1 個 token = 專案名，第 2 個 = 群組名，出現 "SVN" 則 add_svn=True，第二欄 = 工號。
    無法解析則回傳 None。
    """
    line = line.rstrip("\r\n")
    if not line.strip():
        return None

    parts = line.split("\t", 1)
    if len(parts) < 2:
        logger.warning("第 %d 行：缺少 Tab，跳過: %r", line_number, line[:80])
        return None

    first_col = parts[0].strip()
    second_col = parts[1].strip()

    if not first_col or not second_col:
        logger.warning("第 %d 行：第一欄或第二欄為空，跳過", line_number)
        return None

    tokens = first_col.split()
    if len(tokens) < 2:
        logger.warning("第 %d 行：第一欄至少需「專案 群組」兩個 token，跳過: %r", line_number, first_col)
        return None

    if tokens[0].startswith("#"):
        logger.debug("第 %d 行：第一欄為註解，跳過: %r", line_number, first_col[:50])
        return None

    project = tokens[0]
    group = tokens[1]
    add_svn = "SVN" in tokens

    # 工號：第二欄可能有多個欄位（Tab 分隔），取第一個作為工號
    employee_id = second_col.split("\t")[0].strip()
    if not employee_id:
        logger.warning("第 %d 行：工號為空，跳過", line_number)
        return None

    return Record(
        project=project,
        group=group,
        add_svn=add_svn,
        employee_id=employee_id,
        line_number=line_number,
    )


def parse_file(path: str | Path, encoding: str = "utf-8") -> List[Record]:
    """
    讀取文字檔並解析，回傳 Record 列表。空行與無法解析的行寫入 log 並跳過。
    """
    path = Path(path)
    if not path.exists():
        logger.error("檔案不存在: %s", path)
        return []

    records: List[Record] = []
    with open(path, encoding=encoding) as f:
        for line_number, line in enumerate(f, start=1):
            rec = parse_line(line, line_number)
            if rec is not None:
                records.append(rec)

    logger.info("解析完成: %s，共 %d 筆有效紀錄", path, len(records))
    return records
