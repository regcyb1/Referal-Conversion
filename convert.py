"""Batch conversion service and command-line entry point.

With no arguments, open the desktop file chooser. Use --help for CLI usage.
"""

import argparse
from dataclasses import dataclass
import os
from pathlib import Path
import sys
import tempfile
from typing import Callable, Sequence

from pdf_extractor import ConversionError, ExtractedPDF, extract_pdf, safe_sheet_name


@dataclass(frozen=True)
class InputPDF:
    path: Path
    tab_name: str = ""


@dataclass(frozen=True)
class ConversionResult:
    output: Path
    summary_name: str | None
    sheet_names: tuple[str, ...]
    student_count: int


def sheet_names(inputs: Sequence[InputPDF], main_sheet: str | None = None):
    used = set()
    summary = safe_sheet_name(main_sheet, used) if main_sheet and main_sheet.strip() else None
    names = tuple(safe_sheet_name(item.tab_name or item.path.stem, used) for item in inputs)
    return summary, names


# Layout copied from the supplied reference workbooks (e.g. NPT1F2511IT.xlsx).
HEADER_FILL = "FF1F4E78"
FAIL_FILL = "FFFDE9E7"
FAIL_GRADES = {"F", "F+", "F-", "U"}
ASSESSMENT_WIDTHS = {
    "Incourse": 12.57, "Presentation": 16.14, "Group Assignment": 19.86,
    "Practical Test-1": 19.86, "Practical Test-2": 19.86, "Individual Assignment": 24.29,
}


def _assessment_width(header: str) -> float:
    name, _, weight = header.rpartition(" (")
    width = ASSESSMENT_WIDTHS.get(name, 12.0)
    return width + 0.85 if weight.startswith("100") and width > 12 else width


def _write_workbook(documents: Sequence[ExtractedPDF], destination, main_sheet, names):
    # Deliberately imported only at export: PDF inspection also works independently.
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    header_fill = PatternFill("solid", fgColor=HEADER_FILL)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_font = Font(name="Calibri", size=11)
    center = Alignment(horizontal="center")
    fail_fill = PatternFill("solid", fgColor=FAIL_FILL)

    def write_header(sheet, headers):
        sheet.append(list(headers))
        for cell in sheet[1]:
            cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, header_align, border
        sheet.row_dimensions[1].height = 31.5
        sheet.freeze_panes = "A2"
        sheet.page_setup.orientation = "landscape"

    workbook = Workbook()
    workbook.remove(workbook.active)
    if main_sheet:
        index = workbook.create_sheet(main_sheet)
        write_header(index, ["S.No", "Tab", "Source PDF", "Students"])
        for number, (document, name) in enumerate(zip(documents, names), 1):
            index.append([number, name, document.source.name, len(document.rows)])
            row = index.max_row
            index.cell(row, 2).hyperlink = f"#'{name.replace(chr(39), chr(39) * 2)}'!A1"
            for col in range(1, 5):
                cell = index.cell(row, col)
                cell.font, cell.border = body_font, border
                if col in (1, 4):
                    cell.alignment = center
        for col, width in enumerate([7, 34, 60, 10], 1):
            index.column_dimensions[get_column_letter(col)].width = width

    for document, name in zip(documents, names):
        sheet = workbook.create_sheet(name)
        headers = document.headers
        last = len(headers)
        count = last - 8  # assessment columns start at E
        write_header(sheet, headers)
        for excel_row, student in enumerate(document.rows, 2):
            sheet.append([None if value == "" else value for value in student.values])
            failed = str(student.values[last - 2]).strip() in FAIL_GRADES
            for col in range(1, last + 1):
                cell = sheet.cell(excel_row, col)
                cell.font, cell.border = body_font, border
                if col not in (2, last):  # Name and Remarks stay left-aligned
                    cell.alignment = center
                if failed:
                    cell.fill = fail_fill
            sheet.cell(excel_row, 6 + count).number_format = "0.0"
            sheet.row_dimensions[excel_row].height = 15.75
            for col, marker in enumerate(student.flags, 5):
                if marker:
                    sheet.cell(excel_row, col).comment = Comment(f"PDF marker: {marker}", "converter")
        widths = [7, 32, 13, 15] + [_assessment_width(h) for h in headers[4:4 + count]] + [12, 12, 8, 16]
        for col, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(col)].width = width
    workbook.save(destination)
    workbook.close()


def convert_pdfs(
    inputs: Sequence[InputPDF], output: str | Path, main_sheet: str | None = None,
    *, overwrite: bool = False, progress: Callable[[str], None] | None = None,
) -> ConversionResult:
    if not inputs:
        raise ConversionError("Select at least one PDF")
    destination = Path(output).expanduser().resolve()
    if destination.suffix.casefold() != ".xlsx":
        raise ConversionError("The output filename must end in .xlsx")
    if destination.exists() and not overwrite:
        raise ConversionError(f"Output already exists: {destination}. Choose another name or allow replacement.")
    paths = [Path(item.path).expanduser().resolve() for item in inputs]
    if len(set(paths)) != len(paths):
        raise ConversionError("The same PDF was selected more than once")
    for path in paths:
        if path.suffix.casefold() != ".pdf" or not path.is_file():
            raise ConversionError(f"Not a readable PDF file: {path}")
    summary, names = sheet_names(inputs, main_sheet)
    documents = []
    # Validate the entire batch before creating or replacing the output file.
    for index, path in enumerate(paths, 1):
        if progress:
            progress(f"Reading PDF {index}/{len(paths)}: {path.name}")
        documents.append(extract_pdf(path))
    if progress:
        progress("All PDFs validated. Writing Excel workbook…")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".pdf-excel-", suffix=".xlsx", dir=destination.parent, delete=False) as temporary:
            temporary_path = Path(temporary.name)
        _write_workbook(documents, temporary_path, summary, names)
        # No incomplete workbook, and no silent replacement of an existing file.
        if overwrite:
            os.replace(temporary_path, destination)
        else:
            os.link(temporary_path, destination)
            temporary_path.unlink()
        return ConversionResult(destination, summary, names, sum(len(doc.rows) for doc in documents))
    except FileExistsError as error:
        raise ConversionError(f"Output already exists: {destination}") from error
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        try:
            from desktop import launch
            launch()
        except Exception as error:
            print(
                f"Cannot open the desktop window: {error}\n"
                "Make sure Tkinter is installed and run from a desktop session. "
                "Command-line conversion also works without Tkinter: "
                "python convert.py FOLDER --output results.xlsx",
                file=sys.stderr,
            )
            return 1
        return 0
    parser = argparse.ArgumentParser(description="Convert assessment PDFs to one Excel workbook, one tab per PDF.")
    parser.add_argument("pdfs", type=Path, nargs="+", help="PDF files or folders")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output .xlsx filename")
    parser.add_argument("--main-sheet", default=None, help="Optional: add a first index sheet with this name (off by default, like the reference workbooks)")
    parser.add_argument("--tab", action="append", default=[], help="Tab name in input order; repeat for every PDF (empty = PDF name)")
    parser.add_argument("--recursive", action="store_true", help="Include PDFs in subfolders")
    parser.add_argument("--overwrite", action="store_true", help="Explicitly permit replacing an existing output workbook")
    args = parser.parse_args(argv)
    paths = []
    for path in args.pdfs:
        if path.is_dir():
            iterator = path.rglob("*") if args.recursive else path.iterdir()
            paths.extend(sorted((p for p in iterator if p.is_file() and p.suffix.casefold() == ".pdf"), key=str))
        else:
            paths.append(path)
    if args.tab and len(args.tab) != len(paths):
        parser.error("Provide one --tab per PDF, in input order; use --tab '' to keep a PDF filename")
    try:
        result = convert_pdfs(
            [InputPDF(path, args.tab[i] if args.tab else "") for i, path in enumerate(paths)],
            args.output, args.main_sheet, overwrite=args.overwrite, progress=print,
        )
    except (ConversionError, OSError, ImportError) as error:
        print(f"Conversion failed: {error}", file=sys.stderr)
        return 1
    print(f"Saved {result.output}: {len(result.sheet_names)} PDF tabs, {result.student_count:,} student rows.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
