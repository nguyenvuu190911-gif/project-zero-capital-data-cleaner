"""Read and write value-only, single-sheet XLSX files using Python's standard library."""
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}


def _column_index(ref):
    letters = re.match(r"[A-Z]+", ref)
    if not letters:
        raise ValueError("Invalid cell reference")
    value = 0
    for letter in letters.group():
        value = value * 26 + ord(letter) - 64
    return value - 1


def _column_name(index):
    result = ""
    while index >= 0:
        index, remainder = divmod(index, 26)
        result = chr(65 + remainder) + result
        index -= 1
    return result


def read_xlsx(path: Path):
    try:
        with zipfile.ZipFile(path) as archive:
            if len(archive.infolist()) > 100 or sum(item.file_size for item in archive.infolist()) > 50_000_000:
                raise ValueError("Workbook exceeds safety limits")
            names = set(archive.namelist())
            if any(name.endswith("vbaProject.bin") for name in names):
                raise ValueError("Macro-enabled files are unsupported")
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets = workbook.findall("m:sheets/m:sheet", NS)
            if len(sheets) != 1:
                raise ValueError("Only one-sheet workbooks are supported")
            relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            relation_id = sheets[0].attrib[f"{{{NS['r']}}}id"]
            target = next((item.attrib["Target"] for item in relations
                           if item.attrib.get("Id") == relation_id), None)
            if not target:
                raise ValueError("Worksheet relationship missing")
            sheet_path = "xl/" + target.lstrip("/").removeprefix("xl/")
            shared = []
            if "xl/sharedStrings.xml" in names:
                strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                shared = ["".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t"))
                          for si in strings]
            sheet = ET.fromstring(archive.read(sheet_path))
            rows = []
            for row in sheet.findall("m:sheetData/m:row", NS):
                cells = {}
                for cell in row.findall("m:c", NS):
                    if cell.find("m:f", NS) is not None:
                        raise ValueError("Formula cells are unsupported")
                    index = _column_index(cell.attrib["r"])
                    value = cell.find("m:v", NS)
                    kind = cell.attrib.get("t")
                    if kind == "s":
                        content = shared[int(value.text)] if value is not None else ""
                    elif kind == "inlineStr":
                        content = "".join(t.text or "" for t in cell.iter(f"{{{NS['m']}}}t"))
                    elif kind == "b":
                        content = "TRUE" if value is not None and value.text == "1" else "FALSE"
                    elif kind in (None, "n", "str"):
                        content = value.text or "" if value is not None else ""
                    else:
                        raise ValueError(f"Unsupported cell type: {kind}")
                    cells[index] = content
                rows.append([cells.get(i, "") for i in range(max(cells) + 1)] if cells else [])
            return rows
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise ValueError(f"Invalid XLSX: {exc}") from exc


def write_xlsx(path: Path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    xml_rows = []
    for row_number, row in enumerate(rows, 1):
        cells = []
        for index, value in enumerate(row):
            if value == "":
                continue
            reference = f"{_column_name(index)}{row_number}"
            cells.append(f'<c r="{reference}" t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>')
        xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
    sheet = ('<?xml version="1.0" encoding="UTF-8"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>')
    content_types = ('<?xml version="1.0" encoding="UTF-8"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                     '</Types>')
    relationships = ('<?xml version="1.0" encoding="UTF-8"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                     '</Relationships>')
    workbook = ('<?xml version="1.0" encoding="UTF-8"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Cleaned Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
    workbook_relations = ('<?xml version="1.0" encoding="UTF-8"?>'
                          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                          '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in (("[Content_Types].xml", content_types), ("_rels/.rels", relationships),
                           ("xl/workbook.xml", workbook), ("xl/_rels/workbook.xml.rels", workbook_relations),
                           ("xl/worksheets/sheet1.xml", sheet)):
            archive.writestr(name, data)
