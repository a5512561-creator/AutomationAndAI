# -*- coding: utf-8 -*-
"""
OneNote XML Parser
解析 OneNote 頁面 XML，提取表格結構化資料 (ReportItem)。
"""

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Optional

ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
NS = {"one": ONE_NS}

EXPECTED_HEADERS = ["Priority", "Item", "Prgs %", "start", "end", "BW spent%", "過去一週報告"]

_HTML_TAG_RE = re.compile(r"<[^>]+>")
_PCT_RE = re.compile(r"([\d.]+)\s*%")
_STRIKETHROUGH_RE = re.compile(r"text-decoration:\s*line-through", re.IGNORECASE)


@dataclass
class ReportItem:
    priority: Optional[int]
    item: str
    progress_pct: Optional[float]
    start: str
    end: str
    end_has_strikethrough: bool
    bw_spent_pct: Optional[float]
    weekly_note: str


@dataclass
class PageData:
    title: str
    headers: list[str] = field(default_factory=list)
    items: list[ReportItem] = field(default_factory=list)
    header_match: bool = True


def strip_html_tags(text: str) -> str:
    return _HTML_TAG_RE.sub("", text)


def detect_strikethrough(cdata_text: str) -> bool:
    return bool(_STRIKETHROUGH_RE.search(cdata_text))


def parse_percentage(text: str) -> Optional[float]:
    text = strip_html_tags(text).strip()
    if not text:
        return None
    m = _PCT_RE.search(text)
    if m:
        return float(m.group(1)) / 100.0
    return None


def extract_cell_text(cell_elem) -> str:
    """Recursively extract plain text from a <one:Cell> element."""
    parts = []
    for oe in cell_elem.iter(f"{{{ONE_NS}}}OE"):
        for t_elem in oe.findall(f"{{{ONE_NS}}}T"):
            raw = t_elem.text or ""
            clean = strip_html_tags(raw).replace("\n", " ").replace("\r", " ")
            clean = re.sub(r"<br\s*/?>", "\n", raw, flags=re.IGNORECASE)
            clean = strip_html_tags(clean).strip()
            if clean:
                parts.append(clean)
    return "\n".join(parts)


def extract_cell_raw_cdata(cell_elem) -> str:
    """Extract raw CDATA (with HTML) from all <one:T> in a cell."""
    parts = []
    for t_elem in cell_elem.iter(f"{{{ONE_NS}}}T"):
        raw = t_elem.text or ""
        if raw:
            parts.append(raw)
    return "\n".join(parts)


def extract_table(page_root) -> Optional[ET.Element]:
    """Find the first <one:Table> in the page."""
    for table in page_root.iter(f"{{{ONE_NS}}}Table"):
        return table
    return None


def parse_table_rows(table_elem) -> tuple[list[str], list[list]]:
    """Parse header and data rows from a Table element.
    Returns (headers, rows) where each row is a list of (plain_text, raw_cdata) tuples.
    """
    rows = list(table_elem.findall(f"{{{ONE_NS}}}Row"))
    if not rows:
        return [], []

    has_header = table_elem.get("hasHeaderRow", "false").lower() == "true"

    headers = []
    data_rows = []
    start = 0

    if has_header and rows:
        header_row = rows[0]
        for cell in header_row.findall(f"{{{ONE_NS}}}Cell"):
            headers.append(extract_cell_text(cell).strip())
        start = 1

    for row in rows[start:]:
        cells = row.findall(f"{{{ONE_NS}}}Cell")
        row_data = []
        for cell in cells:
            plain = extract_cell_text(cell)
            raw = extract_cell_raw_cdata(cell)
            row_data.append((plain, raw))
        data_rows.append(row_data)

    return headers, data_rows


def _safe_int(s: str) -> Optional[int]:
    s = strip_html_tags(s).strip()
    try:
        return int(s)
    except (ValueError, TypeError):
        return None


def row_to_report_item(row_data: list[tuple[str, str]], n_cols: int = 7) -> Optional[ReportItem]:
    """Convert a parsed row (list of (plain, raw) tuples) to ReportItem.
    Expects 7 columns in order: Priority, Item, Prgs%, start, end, BW spent%, 過去一週報告.
    Gracefully handles tables with fewer columns by padding with empty values.
    """
    target = max(n_cols, 7)
    if len(row_data) < target:
        padded = row_data + [("", "")] * (target - len(row_data))
    else:
        padded = row_data[:target]

    priority_plain, _ = padded[0]
    item_plain, _ = padded[1]
    prgs_plain, _ = padded[2]
    start_plain, _ = padded[3]
    end_plain, end_raw = padded[4]
    bw_plain, _ = padded[5]
    note_plain, _ = padded[6]

    if not item_plain.strip():
        return None

    return ReportItem(
        priority=_safe_int(priority_plain),
        item=item_plain.strip(),
        progress_pct=parse_percentage(prgs_plain),
        start=start_plain.strip(),
        end=end_plain.strip(),
        end_has_strikethrough=detect_strikethrough(end_raw),
        bw_spent_pct=parse_percentage(bw_plain),
        weekly_note=note_plain.strip(),
    )


def parse_page_xml(xml_str: str, page_title: str = "") -> PageData:
    """Parse a full OneNote page XML string, returning PageData."""
    if not xml_str or not xml_str.strip():
        return PageData(title=page_title)

    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError:
        return PageData(title=page_title)

    if not page_title:
        page_title = root.get("name", "")

    table = extract_table(root)
    if table is None:
        return PageData(title=page_title)

    headers, data_rows = parse_table_rows(table)

    header_match = True
    if headers:
        normalized = [h.strip() for h in headers]
        if len(normalized) != len(EXPECTED_HEADERS):
            header_match = False
        else:
            for got, exp in zip(normalized, EXPECTED_HEADERS):
                if got != exp:
                    header_match = False
                    break

    items = []
    for row in data_rows:
        item = row_to_report_item(row, n_cols=len(headers) if headers else 7)
        if item is not None:
            items.append(item)

    return PageData(
        title=page_title,
        headers=headers,
        items=items,
        header_match=header_match,
    )
