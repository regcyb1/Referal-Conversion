"""Compare every extracted value against the supplied workbooks (read-only).

Uses XML/ZIP directly so this audit does not depend on the Excel writer.
Run: python verify_samples.py
"""

from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from pdf_extractor import extract_pdf

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def read_workbook(path):
    with ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in si.findall(".//s:t", NS)) for si in root]
        relationships = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        }
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        for sheet in workbook.findall("s:sheets/s:sheet", NS):
            target = relationships[sheet.attrib[REL]]
            location = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            root = ET.fromstring(archive.read(location))
            rows = []
            for row in root.findall("s:sheetData/s:row", NS):
                values = {}
                for cell in row:
                    col = re.match(r"[A-Z]+", cell.attrib["r"])[0]
                    value = cell.find("s:v", NS)
                    if cell.attrib.get("t") == "s" and value is not None:
                        result = strings[int(value.text)]
                    elif cell.attrib.get("t") == "inlineStr":
                        result = "".join(t.text or "" for t in cell.findall(".//s:t", NS))
                    elif value is not None and value.text is not None:
                        result = float(value.text) if cell.attrib.get("t") not in ("str", "e") else value.text
                    else:
                        result = ""
                    values[col] = result
                if values.get("A") != "" and values.get("A") is not None:
                    rows.append(values)
            yield sheet.attrib["name"], rows


def main():
    root = Path(__file__).resolve().parent
    total, errors, cells = 0, [], 0
    for folder in sorted(root.iterdir()):
        workbook = folder / f"{folder.name}.xlsx"
        if not folder.is_dir() or not workbook.exists():
            continue
        pdfs = sorted(folder.glob("*.pdf"))
        for name, expected in read_workbook(workbook):
            matches = [p for p in pdfs if p.stem.startswith(name.rstrip("-"))]
            if len(matches) != 1:
                errors.append(f"{name}: could not uniquely match a PDF")
                continue
            result = extract_pdf(matches[0])
            actual = [result.headers] + [row.values for row in result.rows]
            if len(expected) != len(actual):
                errors.append(f"{name}: row count {len(actual)} != sample {len(expected)}")
                continue
            local_errors = 0
            for index, (source, sample) in enumerate(zip(actual, expected), 1):
                for col, value in enumerate(source):
                    letter = chr(65 + col)
                    cells += 1
                    if value != sample.get(letter, ""):
                        local_errors += 1
                        errors.append(f"{name}!{letter}{index}: PDF={value!r}; sample={sample.get(letter, '')!r}")
            total += len(result.rows)
            print(f"{'FAIL' if local_errors else 'PASS'} {matches[0].name}: {len(result.rows)} students")
    if not cells:
        print("No sample folders found (FOLDER/FOLDER.xlsx next to its PDFs). Nothing to compare.")
        return False
    print(f"Compared {cells:,} cells / {total:,} student rows; {len(errors)} differences.")
    for error in errors[:30]:
        print(error)
    return bool(errors)


if __name__ == "__main__":
    sys.exit(main())
