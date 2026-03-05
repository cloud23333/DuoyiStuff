from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {
    "x": MAIN_NS,
    "dr": DOC_REL_NS,
    "pr": PKG_REL_NS,
}

CELL_REF_RE = re.compile(r"^([A-Z]+)")


def read_xlsx_table(path: str | Path, sheet_name: str | None = None) -> tuple[list[str], list[dict[str, str]]]:
    """Read the first sheet (or named sheet) as header + row dicts.

    This reader avoids openpyxl's dimension shortcuts so it can parse files that
    report incorrect dimensions but still contain full sheet data.
    """
    file_path = Path(path)
    with ZipFile(file_path) as zf:
        shared_strings = _read_shared_strings(zf)
        sheet_path = _resolve_sheet_path(zf, sheet_name=sheet_name)
        root = ET.fromstring(zf.read(sheet_path))

    sheet_data = root.find("x:sheetData", NS)
    if sheet_data is None:
        raise ValueError(f"No sheetData found in {file_path}")

    raw_rows: list[list[str]] = []
    for row in sheet_data.findall("x:row", NS):
        cells: dict[int, str] = {}
        for cell in row.findall("x:c", NS):
            ref = cell.attrib.get("r", "")
            col_idx = _col_index_from_ref(ref)
            if col_idx is None:
                continue
            cells[col_idx] = _cell_text(cell, shared_strings).strip()
        if not cells:
            continue
        max_col_idx = max(cells)
        row_values = [""] * (max_col_idx + 1)
        for idx, value in cells.items():
            row_values[idx] = value
        raw_rows.append(row_values)

    if not raw_rows:
        raise ValueError(f"No rows found in {file_path}")

    full_header = raw_rows[0]
    active_indices = [i for i, name in enumerate(full_header) if name and str(name).strip()]
    if not active_indices:
        raise ValueError(f"Header row is empty in {file_path}")

    header = [full_header[i].strip() for i in active_indices]
    rows: list[dict[str, str]] = []
    for row_values in raw_rows[1:]:
        row_data: dict[str, str] = {}
        has_value = False
        for idx, col_name in zip(active_indices, header):
            value = row_values[idx] if idx < len(row_values) else ""
            value = value.strip()
            row_data[col_name] = value
            if value:
                has_value = True
        if has_value:
            rows.append(row_data)

    return header, rows


def _read_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("x:si", NS):
        values.append(_join_text_nodes(si))
    return values


def _resolve_sheet_path(zf: ZipFile, sheet_name: str | None) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = workbook.find("x:sheets", NS)
    if sheets is None:
        raise ValueError("No sheets found in workbook.xml")

    selected_sheet = None
    for sheet in sheets.findall("x:sheet", NS):
        if sheet_name is None:
            selected_sheet = sheet
            break
        if sheet.attrib.get("name") == sheet_name:
            selected_sheet = sheet
            break

    if selected_sheet is None:
        raise ValueError(f"Sheet not found: {sheet_name}")

    relationship_id = selected_sheet.attrib.get(f"{{{DOC_REL_NS}}}id")
    if not relationship_id:
        raise ValueError("Sheet relationship id missing")

    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pr:Relationship", NS):
        if rel.attrib.get("Id") != relationship_id:
            continue
        target = rel.attrib.get("Target", "")
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return target
        return f"xl/{target}"

    raise ValueError(f"Could not resolve sheet relationship id: {relationship_id}")


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return _join_text_nodes(cell)

    value_node = cell.find("x:v", NS)
    if value_node is None:
        return ""

    raw = value_node.text or ""
    if cell_type == "s":
        try:
            idx = int(raw)
        except ValueError:
            return raw
        if 0 <= idx < len(shared_strings):
            return shared_strings[idx]
        return raw

    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"

    return raw


def _col_index_from_ref(ref: str) -> int | None:
    match = CELL_REF_RE.match(ref)
    if not match:
        return None
    letters = match.group(1)
    col_num = 0
    for ch in letters:
        col_num = col_num * 26 + (ord(ch) - ord("A") + 1)
    return col_num - 1


def _join_text_nodes(node: ET.Element) -> str:
    return "".join(text_node.text or "" for text_node in node.findall(".//x:t", NS))
