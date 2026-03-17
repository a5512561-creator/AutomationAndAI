import argparse
import os
import sys
from typing import Any, Dict, Optional

import requests
from dotenv import load_dotenv


load_dotenv()


def build_auth_and_headers() -> Dict[str, Any]:
    headers: Dict[str, str] = {"Accept": "application/json"}
    token = os.getenv("CB_TOKEN")
    username = os.getenv("CB_USERNAME")
    password = os.getenv("CB_PASSWORD")

    if token:
        headers["Authorization"] = f"Bearer {token}"
        return {"headers": headers, "auth": None}

    if not username or not password:
        raise RuntimeError("未設定 CB_USERNAME / CB_PASSWORD 或 CB_TOKEN。")
    return {"headers": headers, "auth": requests.auth.HTTPBasicAuth(username, password)}


def cb_patch_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.patch(url, json=payload, timeout=60, verify=True, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"PATCH {url} failed: {resp.status_code} {resp.text}")
    return resp.json() if resp.text else None


def insert_child(base_url: str, parent_id: int, child_id: int, *, index: int = 0) -> None:
    """
    參考 PTC 文件：PATCH /v3/items/{itemId}/children?mode=INSERT
    Request body:
      { "itemReference": { "id": <childId>, "type": "TrackerItemReference" }, "index": 0 }
    """
    url = f"{base_url}/items/{parent_id}/children?mode=INSERT"
    payload = {"itemReference": {"id": child_id, "type": "TrackerItemReference"}, "index": index}
    cb_patch_json(url, payload)


def main(argv: list[str]) -> None:
    parser = argparse.ArgumentParser(description="依 Word 編號重排 Codebeamer item 階層（保留 itemId）")
    parser.add_argument("--base-url", default=(os.getenv("CB_BASE_URL") or "").strip(), help="Codebeamer REST v3 base url")
    parser.add_argument("--apply", action="store_true", help="實際呼叫 API 進行重排")
    args = parser.parse_args(argv)

    base_url = (args.base_url or "").strip()
    if not base_url:
        raise SystemExit("請在 .env 設定 CB_BASE_URL 或用 --base-url")

    # 你提供的 itemId（此腳本只做重排，不建立新 item）
    ids = {
        "root": 585494,  # PaddyTest（根節點）
        "h1": 585496,  # 1 Introduction
        "h11": 585498,  # 1.1 Abbreviation and anonymous
        "h2": 585500,  # 2 Architecture
        "h21": 585502,  # 2.1 Hardware Part Description
        "hwp1": 585504,  # HWP_1
        "hwp2": 585506,  # HWP_2
        "h3": 585508,  # 3 Macro Architecture
        "h31": 585510,  # 3.1 Module Functional Description
    }

    print("將執行的重排動作：")
    print(f"- move {ids['h1']} under {ids['root']}")
    print(f"- move {ids['h2']} under {ids['root']}")
    print(f"- move {ids['h3']} under {ids['root']}")
    print(f"- move {ids['h11']} under {ids['h1']}")
    print(f"- move {ids['h21']} under {ids['h2']}")
    print(f"- move {ids['hwp1']} under {ids['h21']}")
    print(f"- move {ids['hwp2']} under {ids['h21']}")
    print(f"- move {ids['h31']} under {ids['h3']}")
    print()

    if not args.apply:
        print("（未加 --apply）僅顯示計畫，不會呼叫 API。")
        return

    # 任一步失敗就停止，方便手動清理
    insert_child(base_url, ids["root"], ids["h1"], index=0)
    insert_child(base_url, ids["root"], ids["h2"], index=1)
    insert_child(base_url, ids["root"], ids["h3"], index=2)
    insert_child(base_url, ids["h1"], ids["h11"], index=0)
    insert_child(base_url, ids["h2"], ids["h21"], index=0)
    insert_child(base_url, ids["h21"], ids["hwp1"], index=0)
    insert_child(base_url, ids["h21"], ids["hwp2"], index=1)
    insert_child(base_url, ids["h3"], ids["h31"], index=0)

    print("完成重排。請回到 UI 檢查左側樹狀縮排。")


if __name__ == "__main__":
    main(sys.argv[1:])

