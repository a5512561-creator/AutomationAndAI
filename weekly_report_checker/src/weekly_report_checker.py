# -*- coding: utf-8 -*-
"""
Weekly Report Checker
讀取 Power Automate 產出的 OneNote 頁面 JSON，比對本週與上週週報，產出檢查結果 Excel。
"""

import argparse
import json
import os
import re
from datetime import datetime, timedelta
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
except ImportError:
    raise SystemExit("請先安裝依賴: pip install -r requirements.txt")

# 預設輸出目錄（專案下的 output）
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"
# 日期頁面名稱可能格式
DATE_PATTERNS = [
    re.compile(r"^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})$"),
    re.compile(r"^(\d{4})(\d{2})(\d{2})$"),
]


def parse_date_from_page_name(name):
    """從頁面名稱解析日期，支援 2026/03/05、2026-03-05、20260305。"""
    if not name or not isinstance(name, str):
        return None
    s = name.strip()
    for pat in DATE_PATTERNS:
        m = pat.match(s)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                return datetime(y, mo, d).date()
            except ValueError:
                continue
    return None


def get_week_range(d):
    """給定日期 d，回傳該週的 (週一, 週日)。"""
    # weekday(): Monday=0, Sunday=6
    weekday = d.weekday()
    monday = d - timedelta(days=weekday)
    sunday = monday + timedelta(days=6)
    return monday, sunday


def get_this_week_and_last_week():
    """回傳 (本週一, 本週日), (上週一, 上週日)。"""
    today = datetime.now().date()
    this_monday, this_sunday = get_week_range(today)
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    return (this_monday, this_sunday), (last_monday, last_sunday)


def date_in_range(d, start, end):
    if d is None:
        return False
    return start <= d <= end


def normalize_pages_list(pages):
    """將 API 可能回傳的頁面格式轉成 [(date, content_or_none), ...]。"""
    result = []
    if isinstance(pages, list):
        for p in pages:
            if isinstance(p, str):
                result.append((parse_date_from_page_name(p), None))
            elif isinstance(p, dict):
                title = p.get("title") or p.get("displayName") or p.get("name") or ""
                content = p.get("content") or p.get("body") or None
                result.append((parse_date_from_page_name(title), content))
    return result


def get_member_pages_map(data):
    """
    從 JSON 取得 { 同仁名: [(date, content), ...] }。
    支援格式：
    - { "Paddy": ["2026/03/05", ...], ... }
    - { "members": [ { "同仁名": "Paddy", "頁面日期": ["2026/03/05", ...] }, ... ] }
    - { "members": [ { "同仁名": "Paddy", "頁面日期": [...], "頁面內容": {...} }, ... ] }
    """
    out = {}
    if "members" in data and isinstance(data["members"], list):
        for m in data["members"]:
            if not isinstance(m, dict):
                continue
            name = m.get("同仁名") or m.get("member") or m.get("name") or ""
            if not name:
                continue
            pages = m.get("頁面日期") or m.get("pages") or m.get("pageDates") or []
            contents = m.get("頁面內容") or m.get("contents") or {}
            if isinstance(contents, dict):
                # key 為日期字串時可對應內容
                pass
            normalized = normalize_pages_list(pages)
            # 若某頁有內容則帶入
            if contents and isinstance(contents, dict):
                for i, (dt, _) in enumerate(normalized):
                    if dt:
                        key = dt.strftime("%Y/%m/%d") if hasattr(dt, "strftime") else str(dt)
                        if key in contents:
                            normalized[i] = (dt, contents[key])
            out[name] = [(d, c) for d, c in normalized if d is not None]
        return out
    # 格式: { "Paddy": ["2026/03/05", ...], ... }
    for key, val in data.items():
        if key in ("members", "exportTime", "exportTimeUtc"):
            continue
        if isinstance(val, list):
            out[key] = [(parse_date_from_page_name(str(x)), None) for x in val]
            out[key] = [(d, c) for d, c in out[key] if d is not None]
        elif isinstance(val, dict) and "頁面日期" in val:
            pages = val.get("頁面日期") or []
            out[key] = [(parse_date_from_page_name(str(x)), None) for x in pages]
            out[key] = [(d, c) for d, c in out[key] if d is not None]
        else:
            out[key] = []
    return out


def find_page_in_range(pages_with_dates, start, end):
    """在頁面清單中找落在 [start, end] 的日期，回傳 (date, content) 或 (None, None)。"""
    for d, c in pages_with_dates:
        if date_in_range(d, start, end):
            return (d, c)
    return (None, None)


def simple_content_diff(this_content, last_content):
    """
    簡單比對兩段內容，回傳「無進展」的提示。
    若無內容則回傳 None（表示無法比對）。
    """
    if not this_content or not last_content:
        return None
    # 純文字比對：若本週與上週完全相同，視為無更新
    t = (this_content if isinstance(this_content, str) else str(this_content)).strip()
    l = (last_content if isinstance(last_content, str) else str(last_content)).strip()
    if t == l:
        return "本週內容與上週完全相同，無新進展。"
    # 可擴充：逐段或逐行比對，標出相同段落
    return None


def check_one_member(name, pages_with_dates, this_range, last_range):
    """
    回傳檢查結果字串。
    (本週一, 本週日), (上週一, 上週日)
    """
    this_date, this_content = find_page_in_range(pages_with_dates, this_range[0], this_range[1])
    last_date, last_content = find_page_in_range(pages_with_dates, last_range[0], last_range[1])

    if this_date is None and last_date is None:
        return "本週與上週皆未填寫週報。"
    if this_date is None:
        return "本週未填寫週報（上週有：{}）。".format(last_date.strftime("%Y/%m/%d"))
    if last_date is None:
        return "本週有週報（{}），但上週未填寫。".format(this_date.strftime("%Y/%m/%d"))

    no_progress = simple_content_diff(this_content, last_content)
    if no_progress:
        return "本週（{}）與上週（{}）皆有週報，但{}".format(
            this_date.strftime("%Y/%m/%d"),
            last_date.strftime("%Y/%m/%d"),
            no_progress,
        )
    return "本週（{}）與上週（{}）皆有週報，內容有更新。".format(
        this_date.strftime("%Y/%m/%d"),
        last_date.strftime("%Y/%m/%d"),
    )


def find_latest_json(cwd, output_dir):
    """在目錄中找檔名含 onenote_pages 且副檔名為 .json 的最新檔案。"""
    candidates = []
    for folder in (cwd, output_dir):
        if not folder or not folder.exists():
            continue
        for f in folder.glob("*onenote_pages*.json"):
            try:
                candidates.append((f.stat().st_mtime, f))
            except OSError:
                pass
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def main():
    parser = argparse.ArgumentParser(description="Weekly Report Checker：比對本週與上週週報，產出 Excel。")
    parser.add_argument(
        "json_path",
        nargs="?",
        default=None,
        help="Power Automate 產出的 JSON 檔案路徑；不填則自動尋找目前目錄或 output 下最新的 onenote_pages*.json",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Excel 輸出目錄，預設為專案下的 output",
    )
    args = parser.parse_args()

    json_path = args.json_path
    if not json_path:
        cwd = Path.cwd()
        output_dir = Path(args.output_dir) if args.output_dir else DEFAULT_OUTPUT_DIR
        json_path = find_latest_json(cwd, output_dir)
        if not json_path:
            print("錯誤：未指定 JSON 檔案，且找不到 onenote_pages*.json。")
            return 1
        print("使用 JSON 檔案：", json_path)

    path = Path(json_path)
    if not path.exists():
        print("錯誤：找不到檔案", path)
        return 1

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    member_pages = get_member_pages_map(data)
    if not member_pages:
        print("錯誤：JSON 中未解析出任何同仁的頁面資料，請確認格式。")
        return 1

    this_range, last_range = get_this_week_and_last_week()
    results = []
    for name, pages_with_dates in member_pages.items():
        msg = check_one_member(name, pages_with_dates, this_range, last_range)
        results.append((name, msg))

    # 輸出目錄
    out_dir = Path(args.output_dir) if args.output_dir else DEFAULT_OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_name = "WeeklyReport_Check_{}.xlsx".format(ts)
    excel_path = out_dir / excel_name

    wb = Workbook()
    ws = wb.active
    ws.title = "週報檢查"
    # 標題列
    headers = ["同仁名", "weekly report 檢查結果"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    # 資料列
    for row, (name, msg) in enumerate(results, start=2):
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=msg)
        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
    # 欄寬
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 60
    wb.save(excel_path)
    print("已產出：", excel_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
