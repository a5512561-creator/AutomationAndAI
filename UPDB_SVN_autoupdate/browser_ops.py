# -*- coding: utf-8 -*-
"""
UPDB-manager 之 Playwright 操作：專案 UPDB 頁、變更成員名單、依群組加入成員並儲存。
"""
from __future__ import annotations

import logging
import re
import time
from pathlib import Path

import yaml

logger = logging.getLogger(__name__)

# 群組名稱（輸入檔）→ selectors.yaml 的 key；CTC 對應畫面「Add CTC DTD..」
GROUP_LINK_KEYS = {
    "Analog": "link_add_analog",
    "Digital": "link_add_digital",
    "DV": "link_add_dv",
    "Layout": "link_add_layout",
    "APR": "link_add_apr",
    "CTC": "link_add_ctc_dtd",
    "CTC DTD": "link_add_ctc_dtd",
    "Planner": "link_add_planner",
    "Testing": "link_add_testing",
}


def _load_selectors() -> dict:
    path = Path(__file__).resolve().parent / "selectors.yaml"
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _base_url(base: str) -> str:
    """登入頁 URL，結尾去斜線。"""
    return base.rstrip("/")


def _management_zone_root(base: str) -> str:
    """用於直接開啟 .php 的根路徑（不含 index.php），避免導向 .../index.php/updb_member_list_emp.php 造成重新導向。"""
    u = base.rstrip("/")
    if "/index.php" in u:
        u = u.split("/index.php")[0]
    return u


def go_to_project_updb(page, project: str, base_url: str) -> None:
    """導向專案 UPDB 資訊頁；若被 main_page 中斷則重試，仍失敗則改開「變更成員名單」頁。"""
    root = _management_zone_root(base_url)
    info_url = f"{root}/updb_member_list_emp.php?p={project}"
    edit_url = f"{root}/updb_member_edit_emp.php?p={project}"
    logger.info("導向專案 UPDB 頁: %s", info_url)
    err_msg = ""
    for attempt in range(2):
        try:
            page.goto(info_url, wait_until="domcontentloaded", timeout=60000)
            return
        except Exception as e:
            err_msg = str(e)
            if "interrupted" in err_msg.lower() and "main_page" in err_msg:
                logger.warning("導向被 main_page 中斷，等待 2 秒後重試")
                time.sleep(2)
                continue
            if "redirect" in err_msg.lower() or "ERR_TOO_MANY_REDIRECTS" in err_msg:
                logger.error(
                    "導向時發生重新導向錯誤。請以「--clear-cookies」重新執行，並在完成 OTP 後盡快按 Enter。"
                )
            raise
    # 兩次皆被中斷：改為直接開啟「變更成員名單」頁，略過專案資訊頁
    logger.warning("專案資訊頁無法載入，改為直接開啟變更成員名單頁: %s", edit_url)
    try:
        page.goto(edit_url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e2:
        if "interrupted" in str(e2).lower() and "main_page" in str(e2):
            logger.error("變更成員名單頁亦被 main_page 中斷，請以 --clear-cookies 重試或改由瀏覽器手動進入專案後再執行。")
        raise


def go_to_modify_member_list(page, project: str, base_url: str, selectors: dict) -> None:
    """點「變更成員名單」進入修改成員頁；若已在該頁（URL 含 updb_member_edit_emp.php?p=）則略過。"""
    current = page.url
    if "updb_member_edit_emp.php" in current and f"p={project}" in current:
        logger.info("已在變更成員名單頁，略過")
        return
    link_text = selectors.get("link_modify_member_list", "變更成員名單")
    try:
        link = page.get_by_role("link", name=link_text).first
        link.wait_for(state="visible", timeout=10000)
        link.click()
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        logger.info("已點擊「變更成員名單」")
    except Exception as e:
        logger.warning("依連結文字點擊失敗，改為直接開啟 URL: %s", e)
        url = f"{_management_zone_root(base_url)}/updb_member_edit_emp.php?p={project}"
        page.goto(url, wait_until="domcontentloaded", timeout=60000)


def add_members_to_group(
    page,
    context,
    project: str,
    group: str,
    employee_ids: List[str],
    base_url: str,
    selectors: dict,
    wait_seconds: int = 180,
) -> bool:
    """
    點「Add {Group}..」→ 彈窗填工號（逗號分隔）→ Submit → 等待 wait_seconds → 驗證成員出現。
    回傳是否成功。
    """
    link_key = GROUP_LINK_KEYS.get(group)
    if not link_key:
        logger.error("不支援的群組: %s", group)
        return False
    link_text = selectors.get(link_key, f"Add {group}..")
    ids_str = ", ".join(employee_ids)

    try:
        # 畫面上為 <a onclick="window.open(...)"> 內含 <font><b>Add Digital..</b></font>，無 href，故用「a 含文字」定位
        link = page.locator("a").filter(has_text=link_text).first
        try:
            link.wait_for(state="visible", timeout=10000)
        except Exception:
            pattern = re.compile(r"Add\s+" + re.escape(group) + r"\s*\.*", re.I)
            link = page.locator("a").filter(has_text=pattern).first
            link.wait_for(state="visible", timeout=10000)
        with context.expect_page() as popup_info:
            link.click()
        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        popup.wait_for_timeout(500)

        # 彈窗：輸入框為 <input name="RL6483_digital_add" value="">，無 placeholder，說明文字在上方 td
        input_el = (
            popup.locator("form[name='childform'] input:not([type='submit'])")
            .or_(popup.locator("input[name*='_add']"))
            .or_(popup.get_by_placeholder(selectors.get("popup_input_employee_ids_placeholder", "請填入欲新增至")))
            .or_(popup.locator("input[type='text']"))
        ).first
        input_el.wait_for(state="visible", timeout=10000)
        input_el.fill(ids_str)

        submit_text = selectors.get("button_submit", "Submit")
        popup.get_by_role("button", name=submit_text).or_(
            popup.get_by_role("link", name=submit_text)
        ).first.click()

        logger.info("已送出，等待小視窗自動關閉（最多 %d 秒）...", wait_seconds + 120)
        # 彈窗會自動關閉，改為等待 close 事件；若在彈窗上 wait_for_timeout 會因彈窗先關而 TargetClosedError
        try:
            popup.wait_for_event("close", timeout=(wait_seconds + 120) * 1000)
        except Exception:
            try:
                popup.close()
            except Exception:
                pass
        page.wait_for_load_state("domcontentloaded", timeout=10000)
        # 必須按「儲存變更」才會寫入 UPDB；會跳出 confirm，確定後開 updb_member_edit_process_emp.php 視窗（Stage 2…更新UPDB資料庫），約 3～5 分鐘才關閉
        save_text = selectors.get("link_save_changes", "儲存變更")
        page.once("dialog", lambda dialog: dialog.accept())
        with context.expect_page(timeout=10000) as process_info:
            page.locator("a").filter(has_text=save_text).first.click()
        try:
            process_page = process_info.value
            logger.info("儲存變更處理中（Stage 2 更新 UPDB…），請勿手動關閉視窗，約 3～5 分鐘…")
            process_page.wait_for_event("close", timeout=330000)
        except Exception:
            pass
        page.wait_for_timeout(1000)
        logger.info("Add %s 流程完成，已按儲存變更", group)
        return True
    except Exception as e:
        logger.exception("add_members_to_group 失敗: %s", e)
        return False

