# -*- coding: utf-8 -*-
"""
UPDB-manager 批次加人：讀取文字檔 → 開啟瀏覽器 → 等待登入 → 依序執行 UPDB 加成員。
"""
from __future__ import annotations

import argparse
import logging
import sys
import time
from collections import defaultdict
from pathlib import Path

from parser import Record, parse_file
from browser_ops import (
    _load_selectors,
    go_to_project_updb,
    go_to_modify_member_list,
    add_members_to_group,
    save_member_list_changes,
)


def load_config(config_path: str | Path | None) -> dict:
    """載入 config；若路徑不存在則嘗試 config.yaml / config.example.yaml，最後回傳預設。"""
    import yaml
    base = Path(__file__).resolve().parent
    to_try = []
    if config_path and Path(config_path).exists():
        to_try.append(Path(config_path))
    to_try.extend([base / "config.yaml", base / "config.example.yaml"])
    for path in to_try:
        if path.exists():
            with open(path, encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
                cfg.setdefault("updb_login_url", "https://project.rd.realtek.com/ManagementZone/")
                cfg.setdefault("member_add_wait_seconds", 180)
                return cfg
    return {
        "updb_login_url": "https://project.rd.realtek.com/ManagementZone/",
        "member_add_wait_seconds": 180,
    }


def setup_logging(log_dir: Path | None = None) -> None:
    """設定 logging：console + 可選 log 檔。"""
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    logging.basicConfig(level=logging.INFO, format=fmt)
    root = logging.getLogger()
    if log_dir:
        log_dir.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_dir / "updb_batch.log", encoding="utf-8")
        fh.setFormatter(logging.Formatter(fmt))
        root.addHandler(fh)


def group_records_by_project_group(records: list[Record]) -> list[tuple[str, str, list[str]]]:
    """將紀錄依 (project, group) 分組，合併工號。回傳 [(project, group, [employee_id, ...]), ...]"""
    groups: dict[tuple[str, str], list[Record]] = defaultdict(list)
    for r in records:
        groups[(r.project, r.group)].append(r)
    return [(project, group, [r.employee_id for r in recs]) for (project, group), recs in groups.items()]


def run(
    input_path: str | Path,
    base_url: str,
    wait_seconds: int,
    selectors: dict,
    continue_on_error: bool = True,
    clear_cookies: bool = False,
):
    """主流程：Playwright 有頭模式、開登入頁、等 Enter、依序執行 UPDB 加成員。"""
    from playwright.sync_api import sync_playwright

    records = parse_file(input_path)
    if not records:
        logging.getLogger(__name__).warning("無有效紀錄，結束")
        return

    batches = group_records_by_project_group(records)
    logger = logging.getLogger(__name__)
    logger.info("共 %d 筆紀錄，合併為 %d 個批次", len(records), len(batches))

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        if clear_cookies:
            context.clear_cookies()
            logger.info("已清除 Cookie（使用全新 session）")
        page = context.new_page()
        try:
            page.goto(base_url, wait_until="domcontentloaded", timeout=60000)
            logger.info("請在瀏覽器完成登入（含手機 OTP），完成後回到此終端按 Enter 繼續...")
            input()
            # 登入後若停留在 main_page.php 可能觸發「重新導向次數過多」，改為立即導向第一個專案 UPDB 頁以避開
            time.sleep(1)
            # 依專案分組：同一專案內先加完所有群組再按一次「儲存變更」
            projects_order = list(dict.fromkeys(b[0] for b in batches))
            first_project = projects_order[0]
            logger.info("登入後立即導向專案頁以避開重新導向迴圈: %s", first_project)
            go_to_project_updb(page, first_project, base_url)
            go_to_modify_member_list(page, first_project, base_url, selectors)
            current_project: str | None = first_project
            success_updb = 0
            fail_updb = 0

            for project in projects_order:
                project_batches = [(g, ids) for p, g, ids in batches if p == project]
                try:
                    if current_project != project:
                        go_to_project_updb(page, project, base_url)
                        go_to_modify_member_list(page, project, base_url, selectors)
                        current_project = project

                    for group, employee_ids in project_batches:
                        try:
                            ok = add_members_to_group(
                                page, context, project, group, employee_ids,
                                base_url, selectors, wait_seconds,
                                save_after=False,
                            )
                            if ok:
                                success_updb += 1
                            else:
                                fail_updb += 1
                                if not continue_on_error:
                                    raise RuntimeError(f"UPDB 加人失敗: {project} / {group}")
                        except Exception as e:
                            logger.exception("批次失敗 %s / %s: %s", project, group, e)
                            fail_updb += 1
                            if not continue_on_error:
                                raise

                    # 同一專案所有群組都加完後，按一次「儲存變更」
                    if not save_member_list_changes(page, context, selectors):
                        fail_updb += 1
                        if not continue_on_error:
                            raise RuntimeError(f"UPDB 儲存變更失敗: {project}")
                except Exception as e:
                    logger.exception("專案 %s 處理失敗: %s", project, e)
                    if not continue_on_error:
                        raise

            logger.info("完成。UPDB: 成功 %d, 失敗 %d", success_updb, fail_updb)
        finally:
            logger.info("瀏覽器保留開啟，請手動關閉或按 Enter 關閉...")
            try:
                input()
            except (EOFError, KeyboardInterrupt):
                pass
            browser.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="UPDB-manager 依文字檔批次加入專案成員")
    parser.add_argument("input_file", nargs="?", help="輸入文字檔路徑（若未給則從 config 讀）")
    parser.add_argument("--config", "-c", help="設定檔路徑")
    parser.add_argument("--no-continue", action="store_true", help="遇錯誤即中斷，不繼續下一筆")
    parser.add_argument("--log-dir", default="logs", help="日誌目錄（預設 logs）")
    parser.add_argument("--clear-cookies", action="store_true", help="啟動前清除 Cookie，可避免「重新導向次數過多」")
    args = parser.parse_args()

    config_path = args.config or Path(__file__).resolve().parent / "config.yaml"
    if not Path(config_path).exists():
        config_path = Path(__file__).resolve().parent / "config.example.yaml"
    cfg = load_config(config_path)
    base_url = cfg["updb_login_url"].rstrip("/")
    wait_seconds = int(cfg.get("member_add_wait_seconds", 180))
    input_path = args.input_file or cfg.get("input_file")
    if not input_path:
        print("請指定輸入文字檔：main.py <input_file> 或在 config 中設定 input_file", file=sys.stderr)
        sys.exit(1)

    log_dir = Path(args.log_dir)
    setup_logging(log_dir)
    selectors = _load_selectors()

    run(
        input_path=input_path,
        base_url=base_url,
        wait_seconds=wait_seconds,
        selectors=selectors,
        continue_on_error=not args.no_continue,
        clear_cookies=args.clear_cookies,
    )


if __name__ == "__main__":
    main()
