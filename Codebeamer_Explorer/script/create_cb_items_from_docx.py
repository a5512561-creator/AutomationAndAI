import argparse
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


@dataclass
class Node:
    name: str
    category: str
    children: List["Node"] = field(default_factory=list)


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


def extract_images_by_numbering(docx_path: str) -> Dict[str, List[ImageBlob]]:
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
        if not text:
            continue

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

        # 抽圖：抓所有 a:blip 的 r:embed
        try:
            blips = block._p.xpath(".//a:blip", namespaces=namespaces)  # noqa: SLF001
        except Exception:
            blips = []

        if not blips:
            continue

        for idx, blip in enumerate(blips, start=1):
            rel_id = blip.get(f"{{{namespaces['r']}}}embed")
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
    print(f"{indent}- {node.name}  [Category={node.category}]")
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
    resp = requests.get(url, timeout=60, verify=True, **auth_kwargs)
    if resp.status_code >= 400:
        raise RuntimeError(f"GET {url} failed: {resp.status_code} {resp.text}")
    return resp.json()


def cb_post_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.post(url, json=payload, timeout=60, verify=True, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"POST {url} failed: {resp.status_code} {resp.text}")
    return resp.json()


def cb_put_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.put(url, json=payload, timeout=60, verify=True, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"PUT {url} failed: {resp.status_code} {resp.text}")
    return resp.json()


def cb_patch_json(url: str, payload: Dict[str, Any]) -> Any:
    auth_kwargs = build_auth_and_headers()
    headers = dict(auth_kwargs["headers"])
    headers["Content-Type"] = "application/json"
    resp = requests.patch(url, json=payload, timeout=60, verify=True, headers=headers, auth=auth_kwargs["auth"])
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
    resp = requests.post(url, files=files, timeout=120, verify=True, headers=headers, auth=auth_kwargs["auth"])
    if resp.status_code >= 400:
        raise RuntimeError(f"POST {url} failed: {resp.status_code} {resp.text}")
    return resp.json() if resp.text else None


def upload_attachment_v2(rest_root: str, item_id: int, img: ImageBlob) -> Any:
    """
    依 PTC 文件：POST /v2/item/{trackerItemId}/attachment (multipart/form-data, key=attachments)
    """
    url = f"{rest_root}/v2/item/{item_id}/attachment"
    return cb_post_multipart(url, files=[("attachments", (img.filename, img.data, img.content_type))])


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
    """
    將 description 設為 Wiki 文字。
    - 以 tracker 欄位定義取得 description/descriptionFormat 的 fieldId 與 valueModel。
    - 若找不到 descriptionFormat，仍會嘗試只更新 description。
    """
    desc_field_id = find_tracker_field_id_by_tracker_item_field(base_url, tracker_id, "description")
    if desc_field_id is None:
        raise RuntimeError("找不到 description 欄位（trackerItemField=description）")
    desc_def = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields/{desc_field_id}")
    desc_value_model = (desc_def.get("valueModel") or "WikiTextFieldValue").strip()
    desc_name = (desc_def.get("name") or "Description").strip()

    field_values: List[Dict[str, Any]] = [
        {"fieldId": desc_field_id, "type": desc_value_model, "name": desc_name, "value": wiki_text}
    ]

    fmt_field_id = find_tracker_field_id_by_tracker_item_field(base_url, tracker_id, "descriptionFormat")
    if fmt_field_id is not None:
        fmt_def = cb_get_json(f"{base_url}/trackers/{tracker_id}/fields/{fmt_field_id}")
        fmt_value_model = (fmt_def.get("valueModel") or "TextFieldValue").strip()
        fmt_name = (fmt_def.get("name") or "Description Format").strip()
        # 依官方文件慣例：W = Wiki
        field_values.append({"fieldId": fmt_field_id, "type": fmt_value_model, "name": fmt_name, "value": "W"})

    update_item_fields(base_url, item_id, field_values)

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


def create_item_in_tracker(base_url: str, tracker_id: int, name: str, *, parent_id: Optional[int] = None) -> int:
    payload: Dict[str, Any] = {"name": name}
    if parent_id is not None:
        # 使用建立時指定 parent 來建立樹狀（避免某些版本不支援用 /fields 更新 Parent）
        payload["parent"] = {"id": parent_id, "type": "TrackerItemReference"}
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
        item_id = create_item_in_tracker(base_url, tracker_id, node.name, parent_id=parent_id)
        created[id(node)] = item_id
        field_values: List[Dict[str, Any]] = [build_choice_field_value(category_field_id, _category_option_id(node.category))]
        update_item_fields(base_url, item_id, field_values)

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

    # 圖片：將指定章節（例如 1.1）的圖片上傳成附件並插入 description
    if images_by_numbering:
        rest_root = get_rest_root_from_v3_base(base_url)

        def _walk(node: Node) -> Iterator[Node]:
            yield node
            for ch in node.children:
                yield from _walk(ch)

        # 以 node.name 前綴（例如 "1.1 "）辨識 numbering
        for node in _walk(tree):
            m = re.match(r"^(\d+(?:\.\d+)*)\s+", node.name)
            if not m:
                continue
            numbering = m.group(1)
            imgs = images_by_numbering.get(numbering)
            if not imgs:
                continue
            item_id = created[id(node)]
            # 先上傳全部圖片
            for img in imgs:
                upload_attachment_v2(rest_root, item_id, img)
            # 用 wiki markup 插入（附件引用）
            wiki = "\n".join([f"[!{img.filename}!]" for img in imgs])
            update_item_description_wiki(base_url, tracker_id, item_id, wiki)

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

    _, tree = parse_docx_to_tree(docx_path, debug_docx=args.debug_docx)
    images_by_numbering = None if args.no_images else extract_images_by_numbering(docx_path)

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


if __name__ == "__main__":
    main(sys.argv[1:])

