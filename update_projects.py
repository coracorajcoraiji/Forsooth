from __future__ import annotations

import json
import re
import zipfile
from collections import OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "projects.xlsx"
HTML = ROOT / "index.html"

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

CITY_NAMES = {
    "keelung": "基隆市",
    "taipei": "台北市",
    "newtaipei": "新北市",
    "taoyuan": "桃園市",
    "hsinchu": "新竹市",
    "miaoli": "苗栗縣",
    "taichung": "台中市",
    "taichung/changhua": "台中市",
    "changhua": "彰化縣",
    "nantou": "南投縣",
    "yunlin": "雲林縣",
    "chiayi": "嘉義縣",
    "tainan": "台南市",
    "kaohsiung": "高雄市",
    "taitung": "台東市",
}


def read_xml(xlsx: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(xlsx.read(name))


def text_of(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return "".join(element.itertext()).strip()


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    total = 0
    for ch in letters:
        total = total * 26 + (ord(ch.upper()) - ord("A") + 1)
    return total - 1


def shared_strings(xlsx: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in xlsx.namelist():
        return []
    root = read_xml(xlsx, "xl/sharedStrings.xml")
    return [text_of(item) for item in root.findall("main:si", NS)]


def first_sheet_path(xlsx: zipfile.ZipFile) -> str:
    workbook = read_xml(xlsx, "xl/workbook.xml")
    rels = read_xml(xlsx, "xl/_rels/workbook.xml.rels")

    first_sheet = workbook.find("main:sheets/main:sheet", NS)
    if first_sheet is None:
        raise ValueError("No worksheet found in projects.xlsx")

    rel_id = first_sheet.attrib[f"{{{NS['rel']}}}id"]
    for rel in rels.findall("pkgrel:Relationship", NS):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib["Target"].lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"

    raise ValueError("Could not resolve the first worksheet in projects.xlsx")


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    value_type = cell.attrib.get("t")
    if value_type == "s":
        value = cell.find("main:v", NS)
        return strings[int(text_of(value))] if text_of(value) else ""
    if value_type == "inlineStr":
        return text_of(cell.find("main:is", NS))
    return text_of(cell.find("main:v", NS))


def read_rows() -> list[list[str]]:
    with zipfile.ZipFile(WORKBOOK) as xlsx:
        strings = shared_strings(xlsx)
        sheet = read_xml(xlsx, first_sheet_path(xlsx))

    rows: list[list[str]] = []
    for row in sheet.findall(".//main:sheetData/main:row", NS):
        values: list[str] = []
        for cell in row.findall("main:c", NS):
            ref = cell.attrib.get("r", "")
            idx = column_index(ref)
            while len(values) <= idx:
                values.append("")
            values[idx] = cell_value(cell, strings).strip()
        rows.append(values)
    return rows


def load_projects() -> OrderedDict[str, list[dict[str, str]]]:
    rows = read_rows()
    if not rows:
        raise ValueError("projects.xlsx has no data")

    header = [value.strip().lower() for value in rows[0]]
    required = {"item", "city", "title"}
    missing = required - set(header)
    if missing:
        raise ValueError(f"projects.xlsx is missing columns: {', '.join(sorted(missing))}")

    indexes = {name: header.index(name) for name in required}
    projects: OrderedDict[str, list[dict[str, str]]] = OrderedDict()

    for row_number, row in enumerate(rows[1:], start=2):
        while len(row) <= max(indexes.values()):
            row.append("")
        item = row[indexes["item"]].strip()
        city_code = row[indexes["city"]].strip().lower()
        title = row[indexes["title"]].strip()
        if not item and not city_code and not title:
            continue
        if not item or not city_code or not title:
            raise ValueError(f"Row {row_number} is missing item, city, or title")
        if city_code not in CITY_NAMES:
            raise ValueError(f"Row {row_number} has an unknown city code: {city_code}")

        city = CITY_NAMES[city_code]
        projects.setdefault(city, []).append({"num": item, "name": title})

    return projects


def replace_once(text: str, pattern: str, replacement: str, label: str) -> str:
    updated, count = re.subn(pattern, replacement, text, count=1, flags=re.S)
    if count != 1:
        raise ValueError(f"Could not update {label}")
    return updated


def main() -> None:
    projects = load_projects()
    project_count = sum(len(items) for items in projects.values())
    city_count = len(projects)

    html = HTML.read_text(encoding="utf-8")
    project_json = json.dumps(projects, ensure_ascii=False)

    html = replace_once(
        html,
        r"// ── Project Data(?: \(generated from projects\.xlsx\))? ──\s*const PROJECTS = .*?;\s*const CITY_COORDS =",
        f"// ── Project Data (generated from projects.xlsx) ──\nconst PROJECTS = {project_json};\nconst CITY_COORDS =",
        "PROJECTS",
    )

    html = replace_once(
        html,
        r"全台共 <strong style=\"color:var\(--gold-light\)\">\d+</strong> 項工程實績，遍佈 <strong style=\"color:var\(--gold-light\)\">\d+</strong> 個縣市。",
        f"全台共 <strong style=\"color:var(--gold-light)\">{project_count}</strong> 項工程實績，遍佈 <strong style=\"color:var(--gold-light)\">{city_count}</strong> 個縣市。",
        "project summary",
    )

    HTML.write_text(html, encoding="utf-8", newline="\n")
    print(f"Updated {HTML.name}: {project_count} projects, {city_count} cities")


if __name__ == "__main__":
    main()
