import os
import sys
from typing import Any, Dict, List, Optional, Union

import requests
from dotenv import load_dotenv


# 讀取同目錄或專案根目錄下的 .env 檔
load_dotenv()

BASE_URL = os.getenv("CB_BASE_URL", "https://alm.realtek.com/cb/rest/v3")
# 目標 tracker（不同專案/文件只要改 .env）
TRACKER_ID = int(os.getenv("CB_TRACKER_ID", "0"))
# 分頁大小（掃描整個 tracker itemRefs 時使用）
PAGE_SIZE = int(os.getenv("CB_PAGE_SIZE", "100"))
# 單一 item 測試用：CB_TEST_ITEM_ID（從 .env 讀，沒有就用 0）
TEST_ITEM_ID = int(os.getenv("CB_TEST_ITEM_ID", "0"))
# （可選）用來直接呼叫 list API 的完整網址（含 page/pageSize）
TRACKER_ITEMS_URL_OVERRIDE = os.getenv("CB_TRACKER_ITEMS_URL") or None

# （可選）要比對的 component 名稱；多個名稱可用半形逗號或分號分隔（名稱內含逗號時請改用分號）
# 多個名稱可用半形逗號或分號分隔（名稱內含逗號時請改用分號）
CB_TARGET_COMPONENT_NAMES_RAW = (os.getenv("CB_TARGET_COMPONENT_NAMES") or "").strip()
# （可選）若設為正整數，只印出 tracker 前 N 筆 item 的 id/name，方便複製到 CB_TARGET_COMPONENT_NAMES；不設則照常跑比對
CB_LIST_FIRST_N = int(os.getenv("CB_LIST_FIRST_N", "0") or "0")
# 設為 1 時，只依 CB_TARGET_COMPONENT_NAMES（或程式內建預設）列出符合項目的 id/name，不展開 children
CB_LIST_FIRST_LEVEL_ONLY = (os.getenv("CB_LIST_FIRST_LEVEL_ONLY") or "").strip() in ("1", "true", "yes")

# 預設使用 Basic Auth，帳密從環境變數讀取
USERNAME = os.getenv("CB_USERNAME")
PASSWORD = os.getenv("CB_PASSWORD")

# 若你們是用 API Token，可以改用 CB_TOKEN 並在 build_auth_headers 裏調整
TOKEN = os.getenv("CB_TOKEN")


def build_auth_and_headers() -> Dict[str, Any]:
    """
    建立 requests 需要用的 auth / headers。

    - 若設定了 CB_TOKEN，就使用 Bearer Token 認證。
    - 否則使用 Basic Auth（CB_USERNAME / CB_PASSWORD）。
    """
    headers: Dict[str, str] = {"Accept": "application/json"}
    auth: Optional[requests.auth.AuthBase] = None

    if TOKEN:
        headers["Authorization"] = f"Bearer {TOKEN}"
    else:
        if not USERNAME or not PASSWORD:
            raise RuntimeError(
                "未找到 CB_USERNAME / CB_PASSWORD 或 CB_TOKEN，"
                "請先在系統環境變數或 PowerShell 中設定。"
            )
        auth = requests.auth.HTTPBasicAuth(USERNAME, PASSWORD)

    return {"headers": headers, "auth": auth}


def fetch_single_item(item_id: int, verify_ssl: bool = True) -> Dict[str, Any]:
    """
    透過 /v3/items/{itemId} 讀取單一 tracker item。
    """
    if not item_id:
        raise SystemExit("請先在 .env 設定 CB_TEST_ITEM_ID，例如 CB_TEST_ITEM_ID=153320")

    endpoint = f"{BASE_URL}/items/{item_id}"

    auth_kwargs = build_auth_and_headers()

    try:
        resp = requests.get(
            endpoint,
            timeout=30,
            verify=verify_ssl,
            **auth_kwargs,
        )
    except requests.exceptions.RequestException as exc:
        raise SystemExit(f"呼叫 Codebeamer 失敗：{exc}") from exc

    if resp.status_code >= 400:
        raise SystemExit(
            f"Codebeamer 回傳錯誤狀態碼 {resp.status_code}: {resp.text}"
        )

    try:
        data = resp.json()
        if not isinstance(data, dict):
            raise SystemExit(
                f"預期回傳單一物件 dict，但收到：{type(data)}，內容：{data}"
            )
        return data
    except ValueError as exc:
        raise SystemExit(f"回傳內容不是合法 JSON：{exc}\n原始內容：{resp.text}") from exc


def normalize_items_payload(
    payload: Union[List[Dict[str, Any]], Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    目前暫時不用（保留將來要一次抓多筆 items 時），
    先專注在 /items/{itemId} 單筆讀取的情境。
    """
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        return [payload]
    return []


def extract_title(item: Dict[str, Any]) -> str:
    """
    嘗試從 work item JSON 裏取得標題文字。

    常見欄位名稱：
    - name
    - title
    - label
    如有客製欄位，可在此擴充。
    """
    for key in ("name", "title", "label"):
        if key in item and isinstance(item[key], str):
            return item[key]
    # 若完全找不到，回傳空字串避免中斷
    return ""


def debug_single_item(item_id: int) -> None:
    """
    讀取單一 item，印出主要欄位，幫助我們理解 JSON 結構。
    """
    item = fetch_single_item(item_id)
    print("=== 原始 JSON（前幾個欄位） ===")
    for k in list(item.keys())[:20]:
        print(f"{k}: {item[k]}")
    print()

    title = extract_title(item)
    print(f"推測標題欄位內容：{title}")


def fetch_tracker_items_from_url(url: str, verify_ssl: bool = True) -> Dict[str, Any]:
    """
    直接使用完整 URL 讀取 tracker items 分頁結果。
    回傳應該是類似：{"page": ..., "pageSize": ..., "total": ..., "itemRefs": [...]}
    """
    auth_kwargs = build_auth_and_headers()

    try:
        resp = requests.get(
            url,
            timeout=30,
            verify=verify_ssl,
            **auth_kwargs,
        )
    except requests.exceptions.RequestException as exc:
        raise SystemExit(f"呼叫 Codebeamer 失敗：{exc}") from exc

    if resp.status_code >= 400:
        raise SystemExit(
            f"Codebeamer 回傳錯誤狀態碼 {resp.status_code}: {resp.text}"
        )

    try:
        data = resp.json()
        if not isinstance(data, dict):
            raise SystemExit(
                f"預期回傳 dict，但收到：{type(data)}，內容：{data}"
            )
        return data
    except ValueError as exc:
        raise SystemExit(f"回傳內容不是合法 JSON：{exc}\n原始內容：{resp.text}") from exc


def iter_all_tracker_items(tracker_items_base_url: str, page_size: int) -> List[Dict[str, Any]]:
    """
    走訪整個 tracker 的所有 itemRefs，回傳完整清單。

    tracker_items_base_url 不含 page/pageSize 參數，例如：
    https://alm.realtek.com/cb/rest/v3/trackers/206624/items
    """
    all_items: List[Dict[str, Any]] = []
    page = 1

    while True:
        url = f"{tracker_items_base_url}?page={page}&pageSize={page_size}"
        data = fetch_tracker_items_from_url(url)

        items = data.get("itemRefs", [])
        total = data.get("total", 0)

        if not items:
            break

        all_items.extend(items)

        # 若已經超過或等於 total，就可以停止
        if len(all_items) >= total:
            break

        page += 1

    return all_items


def list_top_level_components() -> None:
    """
    掃描整個 tracker items，找出名稱中包含
    '[Template] HW Component Name (Design Doc.)'、'[SWITCH] top view' 等的項目。
    """
    if not TRACKER_ID and not TRACKER_ITEMS_URL_OVERRIDE:
        raise SystemExit("請先在 .env 設定 CB_TRACKER_ID（或提供 CB_TRACKER_ITEMS_URL 覆蓋）。")

    if TRACKER_ITEMS_URL_OVERRIDE:
        tracker_items_base_url = TRACKER_ITEMS_URL_OVERRIDE.split("?", 1)[0]
        page_size = PAGE_SIZE
    else:
        tracker_items_base_url = f"{BASE_URL}/trackers/{TRACKER_ID}/items"
        page_size = PAGE_SIZE

    items = iter_all_tracker_items(tracker_items_base_url, page_size)

    print(f"整個 tracker 共取得 {len(items)} 筆 itemRefs。\n")

    if CB_LIST_FIRST_N > 0:
        n = min(CB_LIST_FIRST_N, len(items))
        print(f"（CB_LIST_FIRST_N={CB_LIST_FIRST_N}，僅列出前 {n} 筆名稱供複製）\n")
        for ref in items[:n]:
            print(f"- [{ref.get('id')}] {ref.get('name')}")
        print("\n可將上列 name 複製到 .env 的 CB_TARGET_COMPONENT_NAMES（多個用分號 ; 分隔）。")
        return

    # 要比對的 component 名稱：優先用 .env 的 CB_TARGET_COMPONENT_NAMES，否則用固定預設（formal）
    if CB_TARGET_COMPONENT_NAMES_RAW:
        sep = ";" if ";" in CB_TARGET_COMPONENT_NAMES_RAW else ","
        target_keywords = [s.strip() for s in CB_TARGET_COMPONENT_NAMES_RAW.split(sep) if s.strip()]
    else:
        target_keywords = [
            "[Template] HW Component Name (Design Doc.)",
            "[SWITCH] top view",
        ]

    if not target_keywords:
        print("未設定要比對的 component 名稱（請設 CB_TARGET_COMPONENT_NAMES）。\n")
        return

    print(f"要比對的 component: {target_keywords}\n")

    def _norm(s: str) -> str:
        return (s or "").replace("\n", " ").replace("\r", " ").strip()

    found: Dict[str, Dict[str, Any]] = {}

    for ref in items:
        name = ref.get("name") or ""
        name_norm = _norm(name)
        for key in target_keywords:
            if name_norm == _norm(key):
                found[key] = ref

    for key in target_keywords:
        ref = found.get(key)
        if ref:
            print(f"[OK] 找到 component：{key} (id={ref.get('id')})")
        else:
            print(f"[MISS] 未找到 component：{key}")
            # 嘗試從 API 回傳中找「名稱包含關鍵字」的項目，供使用者複製正確名稱到 .env
            search_term = ""
            if "[" in key and "]" in key:
                search_term = key[key.index("[") + 1 : key.index("]")]
            else:
                parts = key.strip().split()
                search_term = parts[0] if parts else key[:30]
            if search_term:
                for r in items:
                    nm = (r.get("name") or "")
                    if search_term in nm:
                        print(f"    -> 建議 API 名稱：{_norm(nm)}")
                        break

    print()

    if CB_LIST_FIRST_LEVEL_ONLY:
        print("=== 第一層項目（不展開） ===\n")
        for key in target_keywords:
            ref = found.get(key)
            if ref:
                print(f"- [{ref.get('id')}] {_norm(ref.get('name') or '')}")
        print()
        return

    # 若有找到，進一步列出每個 component 的 children 章節
    for key in target_keywords:
        ref = found.get(key)
        if not ref:
            continue

        item_id = ref["id"]
        item = fetch_single_item(item_id)
        children = item.get("children", [])

        print(f"=== {key} 的 children 章節 (itemId={item_id}) ===")
        if not children:
            print("（沒有 children）\n")
            continue

        for child in children:
            cid = child.get("id")
            cname = child.get("name")
            print(f"- [{cid}] {cname}")
        print()


def main(argv: List[str]) -> None:
    print(f"Base URL: {BASE_URL}")
    print(f"測試 itemId: {TEST_ITEM_ID}")
    if TRACKER_ITEMS_URL_OVERRIDE:
        print(f"Tracker items URL override: {TRACKER_ITEMS_URL_OVERRIDE}")
    else:
        print(f"Tracker ID: {TRACKER_ID}")
    print(f"Page size: {PAGE_SIZE}")
    print()

    # 先印出單一 item（確認連線 OK）
    debug_single_item(TEST_ITEM_ID)

    if CB_LIST_FIRST_LEVEL_ONLY:
        print("\n=== 僅列出第一層（依名稱比對，不展開） ===\n")
        list_top_level_components()
        return

    print("\n=== 掃描整個 tracker，尋找指定 component 並列出其 children ===\n")
    list_top_level_components()


if __name__ == "__main__":
    main(sys.argv[1:])

