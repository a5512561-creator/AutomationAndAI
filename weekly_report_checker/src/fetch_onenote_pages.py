# -*- coding: utf-8 -*-
"""
Fetch OneNote Pages
透過 PowerShell 呼叫 OneNote 桌面應用程式的 COM 介面，讀取週報頁面清單
及最新 2 頁的完整內容，產出 JSON 供 weekly_report_checker.py 進行
填寫狀況分析與內容深度比對。

前提：
  - 已安裝 OneNote 桌面版（2016 / 2021 / Microsoft 365），非 Windows 10 UWP 版
  - 目標 Notebook 已在 OneNote 中開啟並同步
  - 已執行過 fix_onenote_typelib.ps1（首次設定）
"""

import argparse
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PS_SCRIPT = Path(__file__).resolve().parent / "Get-OneNotePages.ps1"


def main():
    parser = argparse.ArgumentParser(
        description="從 OneNote 桌面應用讀取週報頁面，產出 JSON（取代 Power Automate）。"
    )
    parser.add_argument(
        "-n", "--notebook",
        default="Switch-DD member weekly",
        help="OneNote Notebook 名稱（預設：Switch-DD member weekly）",
    )
    parser.add_argument(
        "-m", "--member-list",
        default=None,
        help="成員清單路徑（預設：config/member_list.txt）",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="JSON 輸出目錄（預設：output/）",
    )
    args = parser.parse_args()

    if not PS_SCRIPT.exists():
        print(f"錯誤：找不到 PowerShell 腳本 {PS_SCRIPT}")
        return 1

    ps_args = [
        "powershell", "-ExecutionPolicy", "Bypass",
        "-File", str(PS_SCRIPT),
        "-NotebookName", args.notebook,
    ]
    if args.member_list:
        ps_args += ["-MemberListPath", args.member_list]
    if args.output_dir:
        out_dir = Path(args.output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        ps_args += ["-OutputPath", str(out_dir / f"onenote_pages_{ts}.json")]

    result = subprocess.run(ps_args, capture_output=False)
    return result.returncode


if __name__ == "__main__":
    sys.exit(main())
