"""Strict extraction for the supplied Student Assessment Results PDF format.

No OCR or inferred/recalculated grades: uncertain input fails with a useful error.
"""

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import re

import pymupdf

if hasattr(pymupdf, "no_recommend_layout"):
    pymupdf.no_recommend_layout()  # silence an advisory print on every run


class ConversionError(ValueError):
    """The input cannot be converted without potentially losing data."""


@dataclass(frozen=True)
class StudentRow:
    values: tuple
    page: int
    flags: tuple[str, ...]


@dataclass(frozen=True)
class ExtractedPDF:
    source: Path
    headers: tuple[str, ...]
    rows: tuple[StudentRow, ...]
    page_count: int
    metadata: dict[str, str]


STUDENT_ID = re.compile(r"NP\d+", re.IGNORECASE)
ASSESSMENTS = {
    "exam": "Exam",
    "test": "Test",
    "quiz": "Quiz",
    "project": "Project",
    "incourse": "Incourse",
    "presentation": "Presentation",
    "groupassignment": "Group Assignment",
    "individualassignment": "Individual Assignment",
    "practicaltest-1": "Practical Test-1",
    "practicaltest-2": "Practical Test-2",
}


def clean(value) -> str:
    return " ".join((value or "").split())


def assessment_header(value: str) -> tuple[str, float]:
    match = re.fullmatch(r"(.+?)\s*(\d+(?:\.\d+)?)\s*%", clean(value))
    if not match:
        raise ConversionError(f"Cannot read assessment heading: {value!r}")
    raw_name, weight = match.groups()
    key = re.sub(r"\s+", "", raw_name).casefold()
    name = ASSESSMENTS.get(key)
    if name is None:
        raise ConversionError(f"Unrecognized assessment name: {raw_name!r}")
    return f"{name} ({float(weight):g}%)", float(weight)


def number(value: str, maximum: float, context: str):
    text = clean(value)
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        raise ConversionError(f"{context}: expected a number, found {text!r}")
    result = float(text)
    if not 0 <= result <= maximum:
        raise ConversionError(f"{context}: {text} is outside 0–{maximum:g}")
    return int(result) if result.is_integer() else result


def extract_pdf(path: str | Path) -> ExtractedPDF:
    path = Path(path).expanduser().resolve()
    try:
        with pymupdf.open(path) as document:
            if not document.is_pdf:
                raise ConversionError("The selected file is not a PDF")
            if document.needs_pass:
                raise ConversionError("Password-protected PDFs must be unlocked first")
            if not document.page_count:
                raise ConversionError("The PDF has no pages")
            first_text = document[0].get_text(sort=True)
            if "Student Assessment Results by Class Code" not in first_text:
                raise ConversionError(
                    "Unsupported or scanned PDF. Expected a selectable-text "
                    "'Student Assessment Results by Class Code' report. OCR is not included."
                )
            metadata = {}
            for label in ("Class Code", "Class Description", "Lecturer Name", "Module", "Assessment Code"):
                match = re.search(rf"^{label}\s*:[ \t]*([^\r\n]*)", first_text, re.MULTILINE)
                metadata[label] = clean(match.group(1)) if match else ""
            rows = []
            headers = None
            for page_number, page in enumerate(document, 1):
                # Independent of table detection: every printed student ID must be accounted for.
                printed_ids = Counter(
                    word[4].upper() for word in page.get_text("words")
                    if STUDENT_ID.fullmatch(word[4])
                )
                if not printed_ids:
                    text = page.get_text()
                    if "Statistics" in text or re.search(r"\b(?:Min|Avg|Max)\s*:", text):
                        continue
                    raise ConversionError(f"Page {page_number}: no readable students or statistics")
                page_rows = []
                for table in page.find_tables().tables:
                    data = table.extract()
                    if not any(len(r) > 2 and STUDENT_ID.fullmatch(clean(r[2])) for r in data):
                        continue
                    if table.col_count < 10 or (table.col_count - 8) % 2:
                        raise ConversionError(f"Page {page_number}: unsupported table layout")
                    count = (table.col_count - 8) // 2
                    assessments = [assessment_header(data[0][4 + 2 * i]) for i in range(count)]
                    if abs(sum(weight for _, weight in assessments) - 100) > 0.01:
                        raise ConversionError(f"Page {page_number}: assessment weights do not total 100%")
                    current_headers = (
                        "S.No", "Name", "Student ID", "Course",
                        *(heading for heading, _ in assessments),
                        "Total Marks", "Grade Point", "Grade", "Remarks",
                    )
                    if headers is not None and headers != current_headers:
                        raise ConversionError(f"Page {page_number}: headings changed within the PDF")
                    headers = current_headers
                    for row in data:
                        if len(row) <= 2 or not STUDENT_ID.fullmatch(clean(row[2])):
                            continue
                        context = f"Page {page_number}, student {clean(row[2])}"
                        if not re.fullmatch(r"\d+", clean(row[0])):
                            raise ConversionError(f"{context}: missing or invalid serial number")
                        if not clean(row[1]) or not clean(row[3]):
                            raise ConversionError(f"{context}: missing name or course")
                        values = [int(clean(row[0])), clean(row[1]), clean(row[2]).upper(), clean(row[3])]
                        flags = []
                        for i in range(count):
                            values.append(number(row[4 + 2 * i], 100, context))
                            flags.append(clean(row[5 + 2 * i]))
                        end = 4 + 2 * count
                        values.extend((number(row[end], 100, context), number(row[end + 1], 4, context)))
                        grade = clean(row[end + 2])
                        if not re.fullmatch(r"[A-FU][+-]?", grade):
                            raise ConversionError(f"{context}: unrecognized grade {grade!r}")
                        values.extend((grade, clean(row[end + 3])))
                        page_rows.append(StudentRow(tuple(values), page_number, tuple(flags)))
                extracted_ids = Counter(row.values[2] for row in page_rows)
                if printed_ids != extracted_ids:
                    missing = list((printed_ids - extracted_ids).elements())
                    extra = list((extracted_ids - printed_ids).elements())
                    raise ConversionError(
                        f"Page {page_number}: student count does not reconcile "
                        f"(missing: {', '.join(missing) or 'none'}; extra: {', '.join(extra) or 'none'})"
                    )
                rows.extend(page_rows)
            if not rows or headers is None:
                raise ConversionError("No student results found")
            serials = [row.values[0] for row in rows]
            if serials != list(range(1, len(rows) + 1)):
                raise ConversionError("Student serial numbers are missing, duplicated, or out of order")
            if len({row.values[2] for row in rows}) != len(rows):
                raise ConversionError("The PDF contains duplicate student IDs")
            return ExtractedPDF(path, headers, tuple(rows), document.page_count, metadata)
    except ConversionError as error:
        raise ConversionError(f"{path.name}: {error}") from error
    except Exception as error:
        raise ConversionError(f"{path.name}: cannot read PDF ({error})") from error


def safe_sheet_name(requested: str, used: set[str]) -> str:
    """Excel's 31-character limit and case-insensitive uniqueness both apply."""
    base = re.sub(r"[\x00-\x1f\x7f\\/*?:\[\]]", "_", requested).strip().strip("'")
    base = base or "Sheet"
    if base.casefold() == "history":
        base = "History_"
    candidate = base[:31].strip("'")
    suffix_number = 2
    folded = {name.casefold() for name in used}
    while candidate.casefold() in folded:
        suffix = f" ({suffix_number})"
        candidate = base[:31 - len(suffix)].rstrip("'") + suffix
        suffix_number += 1
    used.add(candidate)
    return candidate
