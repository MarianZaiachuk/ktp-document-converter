#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Core parsing and Excel export logic for the KTP Document Converter."""

from __future__ import annotations

import os
import re
import sys
from copy import copy
from dataclasses import dataclass
from pathlib import Path

import docx
import openpyxl

TEMPLATE_FILENAME = "topics-template.xlsx"
DEFAULT_FOLDER_NAME = "Календарне планування"
TEMPLATE_START_ROW = 6
TEMPLATE_PREFORMATTED_ROWS = 100

SEMESTER_2_RE = re.compile(
    r"(?i)(?:(?<!\w)(?:[IІ]{2}|2(?:-?(?:й|ий|го))?)\s*семестр|семестр\s*(?:[IІ]{2}|2)(?!\w))"
)
SEMESTER_1_RE = re.compile(
    r"(?i)(?:(?<!\w)(?:[IІ]|1(?:-?(?:й|ий|го))?)\s*семестр|семестр\s*(?:[IІ]|1)(?!\w))"
)
GRADE_PARAGRAPH_RE = re.compile(
    r"(\d+\s*[-–—/]\s*\d+|\d+)\s*(?:клас[а-яіїєґ]*)",
    re.IGNORECASE,
)
GRADE_FILENAME_RE = re.compile(
    r"(\d+\s*[-–—_]\s*\d+|\d+)\s*(?:клас[а-яіїєґ]*|_клас)",
    re.IGNORECASE,
)
LESSON_NUM_RE = re.compile(r"^\s*(?:№\s*)?(\d{1,3})\.?\s*$")
MODULE_HEADER_RE = re.compile(r"^(?:тема|модуль|розділ)\b", re.IGNORECASE)


@dataclass(frozen=True)
class Lesson:
    """Single parsed lesson entry from a KTP Word document."""

    semester: int
    module_num: int
    module_name: str
    lesson_num: int
    text: str


@dataclass(frozen=True)
class ParsedPlan:
    """Parsed KTP document metadata and lessons."""

    source_path: str
    folder_name: str
    lessons: list[Lesson]

    @property
    def semester1_count(self) -> int:
        return sum(1 for lesson in self.lessons if lesson.semester == 1)

    @property
    def semester2_count(self) -> int:
        return sum(1 for lesson in self.lessons if lesson.semester == 2)

    @property
    def total_count(self) -> int:
        return len(self.lessons)


@dataclass(frozen=True)
class ConversionResult:
    """Result summary for converting one .docx plan into Excel workbook(s)."""

    source_path: str
    output_dir: str
    created_files: list[str]
    lessons_exported: int


def resolve_template_path(template_path: str | Path | None = None) -> Path:
    """Resolve the path to the Excel template workbook."""
    if template_path is not None:
        explicit = Path(template_path)
        if explicit.is_file():
            return explicit
        raise FileNotFoundError(f"Template file not found: {explicit}")

    candidates: list[Path] = [Path.cwd() / TEMPLATE_FILENAME]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "templates" / TEMPLATE_FILENAME)
    candidates.append(Path(__file__).resolve().parent / "templates" / TEMPLATE_FILENAME)

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    searched = ", ".join(str(p) for p in candidates)
    raise FileNotFoundError(f"Excel template '{TEMPLATE_FILENAME}' not found. Searched: {searched}")


def get_template_workbook(template_path: str | Path | None = None) -> openpyxl.Workbook:
    """Load the blank Yedyna Shkola Excel template workbook."""
    resolved = resolve_template_path(template_path)
    return openpyxl.load_workbook(resolved)


def decode_dropped_path(raw: str | bytes) -> str:
    """Decode a drag-and-drop path from windnd safely on Windows with Cyrillic paths."""
    if isinstance(raw, str):
        return raw

    utf8_decoded: str | None = None
    try:
        utf8_decoded = raw.decode("utf-8")
        if os.path.exists(utf8_decoded) or sys.platform != "win32":
            return utf8_decoded
    except UnicodeDecodeError:
        pass

    ansi_encoding = "mbcs" if sys.platform == "win32" else "cp1251"
    for encoding in (ansi_encoding, "cp1251"):
        try:
            decoded = raw.decode(encoding)
            if os.path.exists(decoded):
                return decoded
        except (UnicodeDecodeError, LookupError):
            continue

    if utf8_decoded is not None:
        return utf8_decoded

    return raw.decode(ansi_encoding, errors="replace")


def detect_folder_name(doc: docx.Document, file_path: str | Path) -> str:
    """Extract the target class folder name (e.g. '8-9 клас') from document text or filename."""
    for paragraph in doc.paragraphs[:15]:
        text = paragraph.text.strip()
        if not text:
            continue
        match = GRADE_PARAGRAPH_RE.search(text)
        if match:
            grade = match.group(1).replace(" ", "").replace("/", "-")
            return f"{grade} клас"

    filename = os.path.basename(str(file_path))
    match = GRADE_FILENAME_RE.search(filename)
    if match:
        grade = match.group(1).replace(" ", "").replace("_", "-")
        return f"{grade} клас"

    return DEFAULT_FOLDER_NAME


def detect_semester_change(text: str) -> int | None:
    """Detect whether a table row switches the current semester (supports Latin and Cyrillic I/II)."""
    if SEMESTER_2_RE.search(text):
        return 2
    if SEMESTER_1_RE.search(text):
        return 1
    return None


def extract_planning_from_docx(docx_path: str | Path) -> ParsedPlan:
    """Parse lessons, modules, semesters, and grade folder name from a .docx KTP file."""
    path_str = str(docx_path)
    doc = docx.Document(path_str)
    lessons: list[Lesson] = []

    current_sem = 1
    current_mod_num = 1
    current_mod_name = "Основна тема"
    mod_counter = 0

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            if not cells or not any(cells):
                continue

            unique_cells: list[str] = []
            for cell in cells:
                if not unique_cells or cell != unique_cells[-1]:
                    unique_cells.append(cell)

            full_row_text = " ".join(part for part in unique_cells if part).strip()

            detected_sem = detect_semester_change(full_row_text)
            if detected_sem is not None:
                current_sem = detected_sem

            first_cell = unique_cells[0]
            lesson_match = LESSON_NUM_RE.fullmatch(first_cell)
            lower_row = full_row_text.lower()

            if "варіативний модуль" in lower_row or (
                lesson_match is None and MODULE_HEADER_RE.match(lower_row)
            ):
                mod_counter += 1
                current_mod_num = mod_counter
                current_mod_name = first_cell if first_cell else full_row_text
                continue

            if lesson_match is not None and len(unique_cells) >= 2:
                lesson_text = unique_cells[1].strip()
                if not lesson_text:
                    continue
                lessons.append(
                    Lesson(
                        semester=current_sem,
                        module_num=current_mod_num,
                        module_name=current_mod_name,
                        lesson_num=int(lesson_match.group(1)),
                        text=lesson_text,
                    )
                )

    folder_name = detect_folder_name(doc, path_str)
    return ParsedPlan(source_path=path_str, folder_name=folder_name, lessons=lessons)


def create_excel_for_semester(
    lessons: list[Lesson],
    output_path: str | Path,
    target_sem: int = 1,
    template_path: str | Path | None = None,
) -> int:
    """Populate the Excel template for one semester and save it to output_path."""
    sem_lessons = [lesson for lesson in lessons if lesson.semester == target_sem]
    if not sem_lessons:
        return 0

    wb = get_template_workbook(template_path)
    ws = wb["теми"] if "теми" in wb.sheetnames else wb.active

    last_template_row = TEMPLATE_START_ROW + TEMPLATE_PREFORMATTED_ROWS - 1
    if len(sem_lessons) > TEMPLATE_PREFORMATTED_ROWS:
        for r_extra in range(last_template_row + 1, len(sem_lessons) + TEMPLATE_START_ROW):
            ws.row_dimensions[r_extra].height = ws.row_dimensions[last_template_row].height
            for c_idx in range(1, 11):
                cell_src = ws.cell(last_template_row, c_idx)
                cell_dst = ws.cell(r_extra, c_idx)
                cell_dst.font = copy(cell_src.font)
                cell_dst.fill = copy(cell_src.fill)
                cell_dst.border = copy(cell_src.border)
                cell_dst.alignment = copy(cell_src.alignment)

    for idx, lesson in enumerate(sem_lessons):
        row_idx = idx + TEMPLATE_START_ROW
        ws.cell(row_idx, 1).value = idx + 1
        ws.cell(row_idx, 2).value = lesson.semester
        ws.cell(row_idx, 3).value = None
        ws.cell(row_idx, 4).value = lesson.module_num
        ws.cell(row_idx, 5).value = lesson.module_name
        ws.cell(row_idx, 6).value = None
        ws.cell(row_idx, 7).value = lesson.lesson_num
        ws.cell(row_idx, 8).value = lesson.text
        ws.cell(row_idx, 9).value = None
        ws.cell(row_idx, 10).value = None

    for row_idx in range(len(sem_lessons) + TEMPLATE_START_ROW, ws.max_row + 1):
        for col_idx in range(2, 11):
            ws.cell(row_idx, col_idx).value = None

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return len(sem_lessons)


def convert_docx_to_excel(
    docx_path: str | Path,
    semester_choice: str = "both",
    output_parent_dir: str | Path | None = None,
    template_path: str | Path | None = None,
    parsed_plan: ParsedPlan | None = None,
) -> ConversionResult:
    """Convert one .docx KTP plan into 1st and/or 2nd semester Excel files."""
    if semester_choice not in ("1", "2", "both"):
        raise ValueError(f"Invalid semester_choice: {semester_choice!r}")

    plan = parsed_plan if parsed_plan is not None else extract_planning_from_docx(docx_path)
    src = Path(docx_path).resolve()
    parent_dir = Path(output_parent_dir) if output_parent_dir is not None else src.parent
    target_dir = parent_dir / plan.folder_name

    created_files: list[str] = []
    lessons_exported = 0

    if not plan.lessons:
        return ConversionResult(
            source_path=str(src),
            output_dir=str(target_dir),
            created_files=created_files,
            lessons_exported=0,
        )

    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = src.stem

    semesters = [1, 2] if semester_choice == "both" else [int(semester_choice)]
    for sem in semesters:
        out_file = target_dir / f"{base_name}_{sem}_сем.xlsx"
        count = create_excel_for_semester(
            plan.lessons,
            out_file,
            target_sem=sem,
            template_path=template_path,
        )
        if count > 0:
            created_files.append(str(out_file))
            lessons_exported += count

    return ConversionResult(
        source_path=str(src),
        output_dir=str(target_dir),
        created_files=created_files,
        lessons_exported=lessons_exported,
    )
