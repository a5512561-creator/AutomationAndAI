import argparse
import base64
import json
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, Iterator, List, Optional, Tuple, Union

import requests
from dotenv import load_dotenv

try:
    from docx import Document  # type: ignore
    from docx.document import Document as _DocxDocument  # type: ignore
    from docx.oxml.text.paragraph import CT_P  # type: ignore
    from docx.oxml.table import CT_Tbl  # type: ignore
    from docx.table import Table  # type: ignore
    from docx.text.paragraph import Paragraph  # type: ignore
except Exception as exc:  # pragma: no cover
    raise SystemExit(
        "缺少套件 python-docx。請先執行：pip install -r requirements.txt\n"
        f"原始錯誤：{exc}"
    ) from exc


load_dotenv()

def _env_bool(name: str, default: bool) -> bool:
    v = (os.getenv(name) or "").strip().lower()
    if not v:
        return default
    return v in ("1", "true", "yes", "y", "on")


CB_VERIFY_SSL = _env_bool("CB_VERIFY_SSL", True)


@dataclass
class Node:
    name: str
    category: str
    children: List["Node"] = field(default_factory=list)
    # 章節內的正文文字（由 Word 段落擷取）
    description_lines: List[str] = field(default_factory=list)


@dataclass
class ImageBlob:
    filename: str
    content_type: str
    data: bytes


# 支援 "HWP_1" 與 "HWP 1"（docx 表格常見會用空白）
HWP_TOKEN_RE = re.compile(r"\bHWP(?:_|\s+)(\d+)\b", re.IGNORECASE)
# 允許常見格式：
# - "1 Introduction"
# - "1. Introduction"
# - "2.1 Hardware Part Description"
# - "2.1. Hardware Part Description"
HEADING_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)(?:\.)?\s+(.+?)\s*$")


def get_paragraph_list_ilvl(paragraph: Paragraph) -> Optional[int]:
    """
    若此段落使用 Word 的「清單自動編號」，會有 numPr/ilvl。
    python-docx 的 paragraph.text 通常不包含編號（例如 1. / 1.1.），需從 XML 取得層級。
    """
    p = paragraph._p  # noqa: SLF001 (python-docx internal)
    ppr = p.pPr
    if ppr is None or ppr.numPr is None or ppr.numPr.ilvl is None:
        return None
    try:
        return int(ppr.numPr.ilvl.val)  # type: ignore[attr-defined]
    except Exception:
        return None


def increment_numbering(counters: List[int], level: int) -> str:
    """
    以 level(1-based) 更新 counters 並回傳章節號，例如 1、1.1、2、2.1。
    """
    if level <= 0:
        level = 1
    while len(counters) < level:
        counters.append(0)
    counters[level - 1] += 1
    for i in range(level, len(counters)):
        counters[i] = 0
    parts = [str(c) for c in counters[:level] if c > 0]
    return ".".join(parts)


def _norm_space(s: str) -> str:
    return " ".join((s or "").replace("\r", " ").replace("\n", " ").split()).strip()


def iter_block_items(parent: _DocxDocument) -> Iterator[Union[Paragraph, Table]]:
    """
    依照 Word 文件實際順序，逐一走訪 Paragraph 與 Table。
    參考 python-docx 官方常見寫法（以底層 XML 判斷 CT_P / CT_Tbl）。
    """
    body = parent.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_images_by_numbering(docx_path: str, *, debug_images: bool = False) -> Dict[str, List[ImageBlob]]:
    """
    以章節編號（例如 1.1）為 key，抽出該章節範圍內段落中的圖片。
    - Word 的自動編號不一定出現在 paragraph.text，因此沿用 ilvl 計數方式取得 numbering。
    - 圖片來源：Paragraph 內的 blip r:embed 關聯圖片 part。
    """
    doc = Document(docx_path)
    numbering_counters: List[int] = []
    current_numbering: Optional[str] = None
    images: Dict[str, List[ImageBlob]] = {}

    namespaces = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    def _ext_from_content_type(ct: str) -> str:
        m = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/jpg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tiff",
            "image/webp": ".webp",
        }
        return m.get(ct.lower(), ".bin")

    for block in iter_block_items(doc):
        if not isinstance(block, Paragraph):
            continue
        text = _norm_space(block.text)

        # 先更新「目前所在章節」：只有在段落有文字時才可能是章節標題
        if text:
            ilvl = get_paragraph_list_ilvl(block)
            if ilvl is not None:
                level = ilvl + 1
                current_numbering = increment_numbering(numbering_counters, level)
            else:
                m = HEADING_RE.match(text)
                if m:
                    current_numbering = m.group(1)

        if not current_numbering:
            continue

        # 若圖片是以 paragraph.text 讀不到的形式存在，可從 runs 的 drawing element 判斷
        if debug_images and text and current_numbering:
            # 只在 debug 時做，避免輸出太多
            pass

        # 抽圖：不要依賴 namespace 前綴（不同文件/版本可能不同），用 local-name() 找 blip
        try:
            blips = block._p.xpath(".//*[local-name()='blip']")  # noqa: SLF001
        except Exception:
            blips = []

        # 某些文件圖片會用 v:imagedata（舊格式）嵌入
        if not blips:
            try:
                blips = block._p.xpath(".//*[local-name()='imagedata']")  # noqa: SLF001
            except Exception:
                blips = []

        if not blips:
            continue

        for idx, blip in enumerate(blips, start=1):
            # r:embed / r:id 取法：不依賴 namespace，直接掃 attribute key
            rel_id = None
            for k, v in getattr(blip, "attrib", {}).items():
                if k.endswith("}embed") or k.endswith("}id"):
                    rel_id = v
                    break
            if not rel_id:
                continue
            part = block.part.related_parts.get(rel_id)
            if not part:
                continue
            content_type = getattr(part, "content_type", "application/octet-stream")
            data = getattr(part, "blob", b"")
            if not data:
                continue
            ext = _ext_from_content_type(content_type)
            filename = f"sec_{current_numbering.replace('.', '_')}_img{idx}{ext}"
            images.setdefault(current_numbering, []).append(
                ImageBlob(filename=filename, content_type=content_type, data=data)
            )

    if debug_images:
        # 額外印出 docx 內總共幾張圖片（不管是否被歸類）
        try:
            total_parts = len(getattr(doc.part, "related_parts", {}))
        except Exception:
            total_parts = -1
        total = sum(len(v) for v in images.values())
        by = ", ".join([f"{k}:{len(v)}" for k, v in sorted(images.items())])
        print(f"[DOCX] images extracted: total={total}, by_numbering={{{by}}}, related_parts={total_parts}")

    return images


def extract_component_name_from_filename(docx_path: str) -> str:
    """
    檔名範例：Hardware-Architecture-Design-Documentation (level4)_PaddyTest.docx
    取出最後一段「_PaddyTest」=> PaddyTest
    """
    base = os.path.basename(docx_path)
    stem, _ = os.path.splitext(base)
    if "_" not in stem:
        return stem
    return stem.split("_")[-1].strip() or stem


def parse_docx_to_tree(docx_path: str, *, debug_docx: bool = False) -> Tuple[str, Node]:
    """
    解析 docx 產生樹狀結構（不呼叫 API）。
    - 根節點：Hardware Component（由檔名擷取）
    - 子節點：依標題編號（1 / 1.1 / 2 / 2.1 / 3 / 3.1 ...）建立 Information 節點
    - 在 2.1 區段內，掃描表格文字抓 HWP_x => 建立 Hardware Part 節點（掛在 2.1 節點底下）
    """
    doc = Document(docx_path)

    component_name = extract_component_name_from_filename(docx_path)
    root = Node(name=component_name, category="Hardware Component")

    # 以 heading level 管理 stack：[(level, node)]
    stack: List[Tuple[int, Node]] = [(0, root)]
    numbering_counters: List[int] = []

    in_hw_part_section = False
    hw_part_anchor_level: Optional[int] = None
    hw_part_anchor_node: Optional[Node] = None
    hw_part_target_node: Optional[Node] = None
    hw_parts_found: List[str] = []
    debug_cell_samples: List[str] = []

    def _maybe_close_hw_section(new_heading_level: int) -> None:
        nonlocal in_hw_part_section, hw_part_anchor_level, hw_part_anchor_node
        if not in_hw_part_section:
            return
        if hw_part_anchor_level is None:
            return
        # 遇到同層或更高層 heading => 結束 2.1 區段
        if new_heading_level <= hw_part_anchor_level:
            in_hw_part_section = False
            hw_part_anchor_level = None
            hw_part_anchor_node = None

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = _norm_space(block.text)
            if not text:
                continue

            ilvl = get_paragraph_list_ilvl(block)
            if ilvl is not None:
                level = ilvl + 1
                numbering = increment_numbering(numbering_counters, level)
                title = text
            else:
                # fallback：若有人把編號打在文字內，仍可吃到
                m = HEADING_RE.match(text)
                if not m:
                    # 非標題段落：附加到目前章節（stack 最底層）的 description_lines
                    if stack:
                        stack[-1][1].description_lines.append(text)
                    continue
                numbering = m.group(1)
                title = m.group(2)
                level = numbering.count(".") + 1

            _maybe_close_hw_section(level)

            node = Node(name=f"{numbering} {title}", category="Information")

            while stack and stack[-1][0] >= level:
                stack.pop()
            parent = stack[-1][1] if stack else root
            parent.children.append(node)
            stack.append((level, node))

            if numbering == "2.1":
                in_hw_part_section = True
                hw_part_anchor_level = level
                hw_part_anchor_node = node
                hw_part_target_node = node
                hw_parts_found.clear()

        elif isinstance(block, Table):
            if not in_hw_part_section or not hw_part_anchor_node:
                continue

            # 掃描整張表格所有 cell
            for row in block.rows:
                for cell in row.cells:
                    cell_text_norm = _norm_space(cell.text)
                    if debug_docx and len(debug_cell_samples) < 8 and cell_text_norm:
                        debug_cell_samples.append(cell_text_norm)
                    for n in HWP_TOKEN_RE.findall(cell_text_norm):
                        hw_parts_found.append(f"HWP_{n}")

    if hw_part_target_node is not None:
        print(f"[DOCX] 2.1 table HWP token count: {len(hw_parts_found)}")
        if debug_docx:
            print("[DOCX] 2.1 table sample cell texts:")
            for s in debug_cell_samples:
                print(f"  - {s}")

    # 於 2.1 節點底下新增 HWP_*
    if hw_parts_found and hw_part_target_node:
        seen: set[str] = set()
        for token in hw_parts_found:
            if token in seen:
                continue
            seen.add(token)
            hw_part_target_node.children.append(Node(name=token, category="Hardware Part"))

    return component_name, root


def print_tree(node: Node, indent: str = "") -> None:
    sample_desc = ""
    if getattr(node, "description_lines", None):
        lines = node.description_lines
        if lines:
            sample_desc = f"  [Desc={lines[0][:30]!r}{'...' if len(lines[0]) > 30 else ''}]"
    print(f"{indent}- {node.name}  [Category={node.category}]{sample_desc}")
    for c in node.children:
        print_tree(c, indent + "  ")


def build_auth_and_headers() -> Dict[str, Any]:
    base_headers: Dict[str, str] = {"Accept": "application/json"}
    token = os.getenv("CB_TOKEN")
    username = os.getenv("CB_USERNAME")
    password = os.getenv("CB_PASSWORD")

    if token:
        base_headers["Authorization"] = f"Bearer {token}"
        return {"headers": base_headers, "auth": None}

    if not username or not password:
        raise RuntimeError("未設定 CB_USERNAME / CB_PASSWORD 或 CB_TOKEN。")
    return {"headers": base_headers, "auth": requests.auth.HTTPBasicAuth(username, password)}


def get_rest_root_from_v3_base(base_url: str) -> str:
    """
    CB_BASE_URL 通常為 .../cb/rest/v3
    這裡回傳 .../cb/rest 以便呼叫 v2 attachment API。
    """
    if base_url.endswith("/v3"):
        return base_url[: -len("/v3")]
    return base_url.rsplit("/v3/", 1)[0] if "/v3/" in base_url else base_url


def cb_get_json(url: str) -> Any:
    auth_kwargs = build_auth_and_headers()
    resp = requests.get(url, timeout=60, verify=CB_VERIFY_SSL, **auth_kwargs)
    if resp.status_code >= 400:
        raise RuntimeError(f"GET {url} failed: {resp.status_code} {resp.text}")
    return resp.json()


def cb_post_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.post(url, json=payload, timeout=60, verify=CB_VERIFY_SSL, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"POST {url} failed: {resp.status_code} {resp.text}")
    return resp.json()


def cb_put_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.put(url, json=payload, timeout=60, verify=CB_VERIFY_SSL, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"PUT {url} failed: {resp.status_code} {resp.text}")
    return resp.json()


def cb_patch_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.patch(url, json=payload, timeout=60, verify=CB_VERIFY_SSL, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"PATCH {url} failed: {resp.status_code} {resp.text}")
    return resp.json() if resp.text else None


def insert_child(base_url: str, parent_id: int, child_id: int, *, index: int) -> None:
    """
    強制縮排/順序：將既有 child 插入到指定 parent 的 children list 位置。
    Endpoint: PATCH /v3/items/{parentId}/children?mode=INSERT
    """
    url = f"{base_url}/items/{parent_id}/children?mode=INSERT"
    payload = {"itemReference": {"id": child_id, "type": "TrackerItemReference"}, "index": index}
    cb_patch_json(url, payload)


def cb_post_multipart(url: str, files: List[Tuple[str, Tuple[str, bytes, str]]]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    # requests 會自動處理 multipart boundary，這裡不要手動設 Content-Type
    resp = requests.post(url, files=files, timeout=120, verify=CB_VERIFY_SSL, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"POST {url} failed: {resp.status_code} {resp.text}")
    return resp.json() if resp.text else None



def cb_post_multipart_raw(url: str, files: List[Tuple[str, Tuple[str, bytes, str]]]) -> Tuple[int, str]:
    raise RuntimeError("probe 模式已移除：cb_post_multipart_raw 不再使用")


def cb_put_multipart_raw(url: str, files: List[Tuple[str, Tuple[str, bytes, str]]]) -> Tuple[int, str]:
    raise RuntimeError("probe 模式已移除：cb_put_multipart_raw 不再使用")


def cb_post_json_raw(url: str, payload: Dict[str, Any]) -> Tuple[int, str]:
    raise RuntimeError("probe 模式已移除：cb_post_json_raw 不再使用")


def cb_put_bytes_raw(url: str, data: bytes, *, content_type: str) -> Tuple[int, str]:
    raise RuntimeError("probe 模式已移除：cb_put_bytes_raw 不再使用")


def cb_get_raw(url: str, *, timeout_s: int = 30) -> Tuple[int, str, int]:
    raise RuntimeError("probe 模式已移除：cb_get_raw 不再使用")


def upload_attachment_v2(rest_root: str, item_id: int, img: ImageBlob) -> Any:
    """
    依 PTC 文件：POST /v2/item/{trackerItemId}/attachment (multipart/form-data, key=attachments)
    """
    url = f"{rest_root}/v2/item/{item_id}/attachment"
    return cb_post_multipart(url, files=[("attachments", (img.filename, img.data, img.content_type))])


def get_display_base_from_api_base(base_url: str) -> str:
    """從 API 基底 (…/cb/rest/v3) 推得網頁基底 (…/cb)，供 displayDocument 連結用。"""
    s = base_url.rstrip("/")
    if s.endswith("/v3"):
        s = s[: -len("/v3")]
    if "/rest" in s:
        s = s.split("/rest")[0]
    return s.rstrip("/") or base_url


def get_api_v3_base_from_rest_v3_base(rest_v3_base: str) -> str:
    """
    部分 CB 環境附件 API 走 /api/v3，而非 /cb/rest/v3。
    例：https://host/cb/rest/v3  => https://host/cb/api/v3
    """
    s = rest_v3_base.rstrip("/")
    if "/cb/rest/v3" in s:
        return s.replace("/cb/rest/v3", "/cb/api/v3")
    if s.endswith("/rest/v3"):
        return s[: -len("/rest/v3")] + "/api/v3"
    # fallback：以 host 為基底
    host = s.split("/rest", 1)[0] if "/rest" in s else s
    return host.rstrip("/") + "/api/v3"


def list_item_attachments(base_url: str, rest_root: str, item_id: int) -> List[Dict[str, Any]]:
    """取得項目的附件列表。先試 v3 再試 v2。回傳 [] 表示無法取得或無附件。"""
    api_v3 = get_api_v3_base_from_rest_v3_base(base_url)
    for url in (
        f"{api_v3}/items/{item_id}/attachments",
        f"{base_url}/items/{item_id}/attachments",
        f"{rest_root}/v2/item/{item_id}/attachments",
    ):
        try:
            out = cb_get_json(url)
            if isinstance(out, list):
                return out
            if isinstance(out, dict) and "attachments" in out:
                return out.get("attachments") or []
        except Exception:
            continue
    return []


def upload_attachment(
    base_url: str, rest_root: str, item_id: int, img: ImageBlob
) -> Optional[Tuple[int, str]]:
    """
    上傳單一圖片為項目附件。先試 v3 再試 v2，multipart 欄位名固定為 "attachments"。
    成功回傳 (attachment_id, filename)，失敗回傳 None 並印出最後一次錯誤。
    若 POST 回傳 []，會改以 GET 附件列表依檔名查找剛上傳的附件 id。
    """
    last_error: Optional[str] = None
    r: Any = None
    api_v3 = get_api_v3_base_from_rest_v3_base(base_url)
    v3_urls = [
        f"{api_v3}/items/{item_id}/attachments",
        f"{api_v3}/items/{item_id}/attachments/content",
        f"{base_url}/items/{item_id}/attachments",
        f"{base_url}/items/{item_id}/attachments/content",
    ]
    for url_v3 in v3_urls:
        try:
            r = cb_post_multipart(url_v3, files=[("attachments", (img.filename, img.data, img.content_type))])
            break
        except Exception as e:
            last_error = str(e).strip()
            r = None
    if r is None:
        try:
            r = upload_attachment_v2(rest_root, item_id, img)
        except Exception as e:
            last_error = str(e).strip()
    # 回傳可能為 list（v2 風格）或單一物件
    if isinstance(r, list) and len(r) > 0:
        first = r[0]
        aid = first.get("id") if isinstance(first, dict) else None
        name = (first.get("name") or img.filename) if isinstance(first, dict) else img.filename
        if isinstance(aid, int):
            return (aid, name)
    if isinstance(r, dict):
        aid = r.get("id")
        name = r.get("name") or img.filename
        if isinstance(aid, int):
            return (aid, name)
    # 若 POST 回傳 []，可能上傳成功但 API 未回傳 id → 用 GET 附件列表依檔名找
    if r == [] or (isinstance(r, list) and len(r) == 0):
        atts = list_item_attachments(base_url, rest_root, item_id)
        for a in atts:
            if not isinstance(a, dict):
                continue
            name = (a.get("name") or "").strip() or (a.get("fileName") or "").strip()
            if name == img.filename:
                aid = a.get("id")
                if isinstance(aid, int):
                    print(f"  [附件] POST 回傳 []，已從 GET 附件列表取得 id={aid} ({img.filename})")
                    return (aid, img.filename)
    # 失敗時一律印出，方便 IT 排查
    if last_error:
        print(f"  [附件] 上傳錯誤（供 IT 排查）: {last_error}")
    else:
        summary = str(r)[:300] if r is not None else "無回應"
        print(f"  [附件] 上傳失敗（回應無法解析，供 IT 排查）: {summary}")
    return None


def find_tracker_field_id_by_tracker_item_field(base_url: str, tracker_id: int, tracker_item_field: str) -> Optional[int]:
    fields = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields")
    if not isinstance(fields, list):
        return None
    for f in fields:
        fid = f.get("id")
        if not isinstance(fid, int):
            continue
        definition = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields/{fid}")
        tif = (definition.get("trackerItemField") or "").strip().lower()
        if tif == tracker_item_field.strip().lower():
            return fid
    return None


def update_item_description_wiki(
    base_url: str,
    tracker_id: int,
    item_id: int,
    wiki_text: str,
) -> None:
    """probe 區塊已移除：此函式不再供主流程使用。"""
    raise RuntimeError("probe 模式已移除：update_item_description_wiki 不再使用")


def images_to_cb_image_macros(uploaded: List[Tuple[int, str]], *, task_id: int, width: int = 600, height: int = 400) -> str:
    """
    使用既有 item(7232) 的 markup 方式，產生 Codebeamer Image macro。
    需要 artifact_id（attachment id），否則前端不會渲染圖片。

    例：
      [{Image wiki='[CB:/displayDocument/MKSImg...png?task_id=7232&artifact_id=16929]' width='600' height='400'}]
    """
    parts: List[str] = []
    for att_id, filename in uploaded:
        link = f"[CB:/displayDocument/{filename}?task_id={task_id}&artifact_id={att_id}]"
        parts.append(f"[{{Image wiki='{link}' width='{width}' height='{height}'}}]")
    return "\n\n".join(parts)

def find_tracker_field_ids(base_url: str, tracker_id: int) -> Tuple[int, int]:
    """
    找出 Category 與 Parent 欄位的 fieldId。
    - Category: 欄位 name=Category（不分大小寫）或 trackerItemField=category
    - Parent: 欄位 name=Parent（不分大小寫）或 trackerItemField=parent
    """
    fields = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields")
    if not isinstance(fields, list):
        raise RuntimeError("tracker fields API 回傳非 list")

    category_field_id: Optional[int] = None
    parent_field_id: Optional[int] = None

    for f in fields:
        fid = f.get("id")
        if not isinstance(fid, int):
            continue
        name = (f.get("name") or "").strip().lower()
        if name == "category":
            category_field_id = fid
        if name == "parent":
            parent_field_id = fid

    # 需要更精準的 trackerItemField，必須 fetch field definition
    if category_field_id is None or parent_field_id is None:
        for f in fields:
            fid = f.get("id")
            if not isinstance(fid, int):
                continue
            if category_field_id is not None and parent_field_id is not None:
                break
            definition = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields/{fid}")
            tif = (definition.get("trackerItemField") or "").strip().lower()
            nm = (definition.get("name") or "").strip().lower()
            if category_field_id is None and (tif == "category" or nm == "category"):
                category_field_id = fid
            if parent_field_id is None and (tif == "parent" or nm == "parent"):
                parent_field_id = fid

    if category_field_id is None:
        raise RuntimeError("找不到 Category 欄位（請確認 tracker 有 Category 欄位，或調整程式比對規則）")
    if parent_field_id is None:
        raise RuntimeError("找不到 Parent 欄位（請確認 tracker 支援樹狀 Parent/Child）")

    return category_field_id, parent_field_id


def get_choice_option_id_by_name(base_url: str, tracker_id: int, field_id: int, option_name: str) -> int:
    definition = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields/{field_id}")
    options = definition.get("options", [])
    if not isinstance(options, list):
        options = []
    for opt in options:
        if (opt.get("name") or "").strip() == option_name:
            oid = opt.get("id")
            if isinstance(oid, int):
                return oid
    raise RuntimeError(f"找不到 Category 選項：{option_name}（請確認 tracker 的 Category option 名稱完全一致）")


def get_item(base_url: str, item_id: int) -> Dict[str, Any]:
    """GET 單一 item，用於驗證 description 是否被伺服器儲存。"""
    return cb_get_json(f"{base_url}/items/{item_id}")


def create_item_in_tracker(
    base_url: str,
    tracker_id: int,
    name: str,
    *,
    parent_id: Optional[int] = None,
    description: Optional[str] = None,
    description_format: Optional[str] = None,
) -> int:
    payload: Dict[str, Any] = {"name": name}
    if parent_id is not None:
        # 使用建立時指定 parent 來建立樹狀（避免某些版本不支援用 /fields 更新 Parent）
        payload["parent"] = {"id": parent_id, "type": "TrackerItemReference"}
    if description is not None:
        payload["description"] = description
    if description_format is not None:
        payload["descriptionFormat"] = description_format
    created = cb_post_json(f"{base_url}/trackers/{tracker_id}/items", payload)
    item_id = created.get("id")
    if not isinstance(item_id, int):
        raise RuntimeError(f"建立 item 失敗，回傳沒有 id：{created}")
    return item_id


def update_item_fields(base_url: str, item_id: int, field_values: List[Dict[str, Any]]) -> None:
    cb_put_json(f"{base_url}/items/{item_id}/fields", {"fieldValues": field_values})


def build_choice_field_value(field_id: int, option_id: int, field_name: str = "Category") -> Dict[str, Any]:
    return {
        "fieldId": field_id,
        "type": "ChoiceFieldValue",
        "name": field_name,
        "values": [{"id": option_id, "type": "ChoiceOptionReference"}],
    }


def apply_tree_to_codebeamer(
    base_url: str,
    tracker_id: int,
    tree: Node,
    *,
    force: bool,
    reindent: bool,
    images_by_numbering: Optional[Dict[str, List[ImageBlob]]] = None,
) -> int:
    category_field_id, _parent_field_id = find_tracker_field_ids(base_url, tracker_id)

    cat_hw_component = get_choice_option_id_by_name(base_url, tracker_id, category_field_id, "Hardware Component")
    cat_hw_part = get_choice_option_id_by_name(base_url, tracker_id, category_field_id, "Hardware Part")
    cat_information = get_choice_option_id_by_name(base_url, tracker_id, category_field_id, "Information")

    rest_root = get_rest_root_from_v3_base(base_url)
    display_base = get_display_base_from_api_base(base_url)

    # 上傳圖片到固定 item（host），用其 attachment 的 artifact_id 生成 Image macro。
    # 這樣 description 只需要在 POST 建立時帶入即可，避免 description PUT 403 問題。
    attachment_host_item_id = int(
        os.getenv("CB_ATTACHMENT_HOST_ITEM_ID")
        or os.getenv("CB_TEST_ITEM_ID")
        or "0"
    )
    if attachment_host_item_id <= 0:
        raise SystemExit("請在 .env 設定 CB_ATTACHMENT_HOST_ITEM_ID（建議填 7232）或至少填 CB_TEST_ITEM_ID")

    # 簡化：不做「同名搜尋」避免 endpoint 差異；若要避免重複，使用者先用 UI / 或之後補強 query API
    if not force:
        print("（--force 未指定）提醒：此測試程式不做同名去重，若要允許重複建立請加 --force。")
        raise SystemExit("為避免重複建立，請加 --force 後再執行 --apply。")

    def _category_option_id(cat: str) -> int:
        if cat == "Hardware Component":
            return cat_hw_component
        if cat == "Hardware Part":
            return cat_hw_part
        return cat_information

    created: Dict[int, int] = {}

    def _create_recursive(node: Node, parent_id: Optional[int]) -> int:
        imgs: Optional[List[ImageBlob]] = None
        if images_by_numbering:
            m = re.match(r"^(\d+(?:\.\d+)*)\s+", node.name)
            if m:
                imgs = images_by_numbering.get(m.group(1))

        text_desc = "\n".join(getattr(node, "description_lines", [])).strip()

        desc: Optional[str] = None
        desc_fmt: Optional[str] = None

        # 重要：你們環境對 Description 欄位 PUT 會 403（not writable），因此 description 只能在 POST 建立時帶入。
        # 同時前端 Image macro 需要 artifact_id，因此這裡必須先把圖片上傳到 host item，拿到 artifact_id。
        if imgs:
            uploaded: List[Tuple[int, str]] = []
            print(f"  [圖片] {node.name!r}: 先上傳 {len(imgs)} 張到 host itemId={attachment_host_item_id}…")
            for img in imgs:
                one = upload_attachment(base_url, rest_root, attachment_host_item_id, img)
                if one:
                    uploaded.append(one)
                else:
                    break
            if len(uploaded) == len(imgs):
                desc = images_to_cb_image_macros(uploaded, task_id=attachment_host_item_id)
                desc_fmt = "Wiki"
                print(f"  [圖片] Image macro 產生完成（artifact_id count={len(uploaded)}）")
            else:
                print(f"  [警告] host 上傳成功 {len(uploaded)}/{len(imgs)}，目前仍需 artifact_id 才可渲染；本節點將不寫入圖片 description")
                desc = None
                desc_fmt = None

        # 若有章節文字，把文字和圖片一起塞進 description（description PUT 在此環境會 403，所以只用 POST）
        if text_desc:
            if desc:
                desc = f"{text_desc}\n\n{desc}"
            else:
                desc = text_desc
            # 文字本身與 Wiki macro 都用 Wiki 格式
            desc_fmt = desc_fmt or "Wiki"

        item_id = create_item_in_tracker(
            base_url,
            tracker_id,
            node.name,
            parent_id=parent_id,
            description=desc,
            description_format=desc_fmt,
        )
        created[id(node)] = item_id

        field_values: List[Dict[str, Any]] = [build_choice_field_value(category_field_id, _category_option_id(node.category))]
        update_item_fields(base_url, item_id, field_values)

        # 驗證：description 是否含 displayDocument（image macro）
        if desc:
            try:
                item = get_item(base_url, item_id)
                stored = (item.get("description") or "").strip()
                has_disp = "displayDocument" in stored
                print(f"  [檢查] itemId={item_id} description_len={len(stored)} has_displayDocument={'是' if has_disp else '否'}")
                print(f"    stored_snippet={stored[:200]!r}")
            except Exception as e:
                print(f"  [檢查] 無法 GET 驗證 description：{e}")

        for child in node.children:
            _create_recursive(child, item_id)
        return item_id

    root_id = _create_recursive(tree, None)

    if reindent:
        # 某些 tracker 的 UI/模型對 parent 欄位不敏感，建立後再用 children INSERT 強制縮排與順序。
        def _reindent(node: Node) -> None:
            parent_item_id = created[id(node)]
            for idx, child in enumerate(node.children):
                insert_child(base_url, parent_item_id, created[id(child)], index=idx)
                _reindent(child)

        _reindent(tree)

    return root_id


def main(argv: List[str]) -> None:
    parser = argparse.ArgumentParser(description="由 DOCX 解析並在 Codebeamer 建立樹狀項目（測試程式）")
    parser.add_argument("--docx-path", default=os.getenv("CB_DOCX_PATH") or "", help="docx 路徑（也可用 CB_DOCX_PATH）")
    parser.add_argument("--dry-run", action="store_true", help="只解析並印出將建立的樹狀結構，不呼叫 API")
    parser.add_argument("--debug-docx", action="store_true", help="除錯：額外列出 2.1 區段 table 前幾個 cell 文字")
    parser.add_argument("--apply", action="store_true", help="實際呼叫 API 建立項目（需要 --force）")
    parser.add_argument("--force", action="store_true", help="允許建立（避免重複建立保護）")
    parser.add_argument("--no-reindent", action="store_true", help="建立後不做 children INSERT 重排（預設會重排）")
    parser.add_argument("--no-images", action="store_true", help="不處理 docx 圖片（預設會將章節內圖片上傳並插入 description）")
    parser.add_argument("--debug-images", action="store_true", help="除錯：印出 docx 圖片抽取統計")
    args = parser.parse_args(argv)

    docx_path = args.docx_path.strip().strip('"')
    if not docx_path:
        raise SystemExit("請提供 --docx-path 或設定環境變數 CB_DOCX_PATH")
    if not os.path.exists(docx_path):
        raise SystemExit(f"找不到 docx：{docx_path}")

    base_url = (os.getenv("CB_BASE_URL") or "").strip()
    tracker_id = int(os.getenv("CB_TRACKER_ID", "0") or "0")
    if not base_url or not tracker_id:
        raise SystemExit("請先在 .env 設定 CB_BASE_URL 與 CB_TRACKER_ID")
    # 探針/除錯模式已移除（保留程式碼但強制不執行，避免誤用）
    if False and (args.probe_attachments or args.probe_description or args.probe_existing_attachment):
        item_id = int(args.probe_item_id or 0)
        if item_id <= 0:
            raise SystemExit("請提供 --probe-item-id 或在 .env 設定 CB_TEST_ITEM_ID")
        rest_root = get_rest_root_from_v3_base(base_url)
        api_v3 = get_api_v3_base_from_rest_v3_base(base_url)

        if args.probe_existing_attachment:
            att_id = int(args.probe_attachment_id or 0)
            if att_id <= 0:
                raise SystemExit("請提供 --probe-attachment-id（artifact_id）")
            print(f"=== 探針：既有附件（itemId={item_id}, attachmentId={att_id}）===\n")
            # 1) 列附件
            print("[probe] list_item_attachments()")
            atts = list_item_attachments(base_url, rest_root, item_id)
            print(f"  count={len(atts)}")
            for a in atts[:5]:
                if isinstance(a, dict):
                    print(f"  - id={a.get('id')} name={a.get('name') or a.get('fileName')}")
            # 2) GET attachment metadata/content（嘗試常見路徑）
            candidates = [
                f"{base_url}/attachments/{att_id}",
                f"{base_url}/attachments/{att_id}/content",
                f"{api_v3}/attachments/{att_id}",
                f"{api_v3}/attachments/{att_id}/content",
                f"{rest_root}/v2/attachment/{att_id}",
                f"{rest_root}/v2/attachment/{att_id}/content",
            ]
            print("\n[probe] GET attachment endpoints")
            for u in candidates:
                status, ct, blen = cb_get_raw(u)
                print(f"  - GET {u}\n    status={status} content-type={ct} bytes={blen}")
            print("\n（若能成功取得 content 且 content-type 為 image/png，代表附件讀取 API 可用；接著就只差『上傳 API』與『回傳 id』）\n")

        if args.probe_attachments:
            print(f"=== 探針：附件上傳（itemId={item_id}）===\n")
            png_1x1 = base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6XG7mQAAAAASUVORK5CYII="
            )
            img = ImageBlob(filename="cursor_probe.png", content_type="image/png", data=png_1x1)
            if args.probe_attachments_verbose:
                candidates = [
                    f"{api_v3}/items/{item_id}/attachments",
                    f"{api_v3}/items/{item_id}/attachments/content",
                    f"{base_url}/items/{item_id}/attachments",
                    f"{base_url}/items/{item_id}/attachments/content",
                    f"{rest_root}/v2/item/{item_id}/attachment",
                ]
                for u in candidates:
                    status, txt = cb_post_multipart_raw(u, files=[("attachments", (img.filename, img.data, img.content_type))])
                    snippet = (txt or "").replace("\r", " ").replace("\n", " ")
                    if len(snippet) > 300:
                        snippet = snippet[:300] + "…"
                    print(f"[probe] POST {u}\n  status={status}\n  body={snippet}\n")
            out = upload_attachment(base_url, rest_root, item_id, img)
            print(f"\n探針結果：upload_attachment => {out!r}")
            atts = list_item_attachments(base_url, rest_root, item_id)
            names = []
            for a in atts:
                if isinstance(a, dict):
                    nm = (a.get("name") or a.get("fileName") or "").strip()
                    if nm:
                        names.append(nm)
            print(f"附件列表（檔名）前 20 筆：{names[:20]}")
            print("\n（若 upload 回傳 None 但附件列表出現 cursor_probe.png，表示上傳成功但回應不含 id，需要 IT 提供正確回應格式/endpoint）")

            if args.probe_attachments_deep:
                print("\n=== deep probe: 測 /cb/rest/v3/items/{id}/attachments 不同欄位名 ===\n")
                base_att_url = f"{base_url}/items/{item_id}/attachments"
                before = set(names)
                field_names = ["attachments", "file", "files", "upload", "content"]
                for fn in field_names:
                    status, txt = cb_post_multipart_raw(
                        base_att_url, files=[(fn, (img.filename, img.data, img.content_type))]
                    )
                    snippet = (txt or "").replace("\r", " ").replace("\n", " ")
                    if len(snippet) > 200:
                        snippet = snippet[:200] + "…"
                    after_names = []
                    try:
                        atts2 = list_item_attachments(base_url, rest_root, item_id)
                        for a in atts2:
                            if isinstance(a, dict):
                                nm2 = (a.get("name") or a.get("fileName") or "").strip()
                                if nm2:
                                    after_names.append(nm2)
                    except Exception:
                        pass
                    added = sorted(set(after_names) - before)
                    print(f"[deep] field={fn!r} status={status} body={snippet}")
                    print(f"       added={added[:5]}")
                # 再測 PUT（有些 API 用 PUT）
                for fn in field_names:
                    status, txt = cb_put_multipart_raw(
                        base_att_url, files=[(fn, (img.filename, img.data, img.content_type))]
                    )
                    snippet = (txt or "").replace("\r", " ").replace("\n", " ")
                    if len(snippet) > 200:
                        snippet = snippet[:200] + "…"
                    after_names = []
                    try:
                        atts2 = list_item_attachments(base_url, rest_root, item_id)
                        for a in atts2:
                            if isinstance(a, dict):
                                nm2 = (a.get("name") or a.get("fileName") or "").strip()
                                if nm2:
                                    after_names.append(nm2)
                    except Exception:
                        pass
                    added = sorted(set(after_names) - before)
                    print(f"[deep-put] field={fn!r} status={status} body={snippet}")
                    print(f"           added={added[:5]}")
                print("\n（若任何欄位名能讓 added 出現 cursor_probe.png，就表示上傳欄位名要用那個）")

                print("\n=== deep probe: 嘗試 JSON 建立 attachment，再 PUT /attachments/{id}/content ===\n")
                # 1) POST 建立 attachment metadata（若支援）
                create_payloads = [
                    {"name": img.filename},
                    {"name": img.filename, "fileName": img.filename},
                    {"name": img.filename, "contentType": img.content_type},
                ]
                created_id: Optional[int] = None
                for p in create_payloads:
                    st, tx = cb_post_json_raw(base_att_url, p)
                    snippet2 = (tx or "").replace("\r", " ").replace("\n", " ")
                    if len(snippet2) > 200:
                        snippet2 = snippet2[:200] + "…"
                    print(f"[deep-json] POST {base_att_url} payload={p} => status={st} body={snippet2}")
                    if st in (200, 201) and tx.strip():
                        try:
                            j = json.loads(tx)
                            if isinstance(j, dict) and isinstance(j.get("id"), int):
                                created_id = int(j["id"])
                                break
                            if isinstance(j, list) and j and isinstance(j[0], dict) and isinstance(j[0].get("id"), int):
                                created_id = int(j[0]["id"])
                                break
                        except Exception:
                            pass
                if created_id:
                    put_url = f"{base_url}/attachments/{created_id}/content"
                    stp, txp = cb_put_bytes_raw(put_url, img.data, content_type=img.content_type)
                    snippet3 = (txp or "").replace("\r", " ").replace("\n", " ")
                    if len(snippet3) > 200:
                        snippet3 = snippet3[:200] + "…"
                    print(f"[deep-json] PUT {put_url} => status={stp} body={snippet3}")
                    # 再列一次附件看是否新增
                    atts3 = list_item_attachments(base_url, rest_root, item_id)
                    after3 = []
                    for a in atts3:
                        if isinstance(a, dict):
                            nm3 = (a.get("name") or a.get("fileName") or "").strip()
                            if nm3:
                                after3.append(nm3)
                    added3 = sorted(set(after3) - before)
                    print(f"[deep-json] after added={added3[:10]}")
                else:
                    print("[deep-json] 無法從 POST 回應取得 attachment id（可能此 endpoint 不支援 JSON 建立）")
        if args.probe_description:
            print(f"\n=== 探針：Description 寫入（itemId={item_id}）===\n")
            try:
                update_item_description_wiki(base_url, tracker_id, item_id, "Cursor probe: !cursor_probe.png!")
                print("Description 寫入：成功")
                try:
                    item = get_item(base_url, item_id)
                    stored = (item.get("description") or "").strip()
                    print(f"[probe] GET description length={len(stored)} contains '!cursor_probe.png!'={'!cursor_probe.png!' in stored}")
                except Exception as ge:
                    print(f"[probe] GET 驗證失敗：{ge}")
            except Exception as e:
                print(f"Description 寫入：失敗 => {e}")
        return

    _, tree = parse_docx_to_tree(docx_path, debug_docx=args.debug_docx)
    images_by_numbering = None if args.no_images else extract_images_by_numbering(docx_path, debug_images=args.debug_images)

    print("=== 解析結果（將建立的樹狀結構）===\n")
    print_tree(tree)
    print()

    if args.dry_run and args.apply:
        raise SystemExit("請擇一使用 --dry-run 或 --apply")
    if args.dry_run or not args.apply:
        return

    root_id = apply_tree_to_codebeamer(
        base_url,
        tracker_id,
        tree,
        force=args.force,
        reindent=not args.no_reindent,
        images_by_numbering=images_by_numbering,
    )
    print(f"\n完成。根節點 itemId={root_id}")
    if images_by_numbering and sum(len(v) for v in images_by_numbering.values()) > 0:
        print("  若網頁上該項目 Description 未顯示圖片，請執行：")
        print("    1) --dry-run --debug-images  確認 docx 是否有抽出圖片")
        print("    2) 在 Codebeamer 開啟該項目，檢視 Description 欄位是否含 [{Html ...}] 或 <img>")


if __name__ == "__main__":
    main(sys.argv[1:])

