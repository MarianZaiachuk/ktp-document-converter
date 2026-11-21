#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Unit and integration tests for converter.py."""

from __future__ import annotations

from pathlib import Path

import docx
import openpyxl
import pytest

from converter import (
    DEFAULT_FOLDER_NAME,
    Lesson,
    convert_docx_to_excel,
    create_excel_for_semester,
    decode_dropped_path,
    detect_folder_name,
    detect_semester_change,
    extract_planning_from_docx,
    resolve_template_path,
)


def _create_sample_docx(path: Path, title: str = "Календарне планування") -> Path:
    doc = docx.Document()
    doc.add_paragraph(title)
    table = doc.add_table(rows=0, cols=3)

    # Semester 1 (Cyrillic І)
    r = table.add_row()
    r.cells[0].merge(r.cells[2])
    r.cells[0].text = "І семестр"

    # Header/non-lesson rows that contain digits (regression test for false lesson parsing)
    r = table.add_row()
    r.cells[0].text = "10 клас"
    r.cells[1].text = "2026 рік"
    r.cells[2].text = "Примітка"

    # Module 1
    r = table.add_row()
    r.cells[0].merge(r.cells[2])
    r.cells[0].text = "Тема 1. Основи алгоритмізації"

    r = table.add_row()
    r.cells[0].text = "1"
    r.cells[1].text = "Поняття алгоритму та виконавця"

    r = table.add_row()
    r.cells[0].text = "№ 2."
    r.cells[1].text = "Лінійні алгоритми"

    # Semester 2 (Cyrillic ІІ)
    r = table.add_row()
    r.cells[0].merge(r.cells[2])
    r.cells[0].text = "ІІ семестр"

    # Variative module
    r = table.add_row()
    r.cells[0].merge(r.cells[2])
    r.cells[0].text = "Варіативний модуль: Python"

    r = table.add_row()
    r.cells[0].text = "3"
    r.cells[1].text = "Змінні та типи даних"

    r = table.add_row()
    r.cells[0].text = "4"
    r.cells[1].text = ""  # Empty lesson text should be skipped

    doc.save(path)
    return path


def test_detect_semester_change_supports_cyrillic_and_latin() -> None:
    assert detect_semester_change("І семестр") == 1
    assert detect_semester_change("I семестр") == 1
    assert detect_semester_change("1-й семестр") == 1
    assert detect_semester_change("семестр 1") == 1

    assert detect_semester_change("ІІ семестр") == 2
    assert detect_semester_change("II семестр") == 2
    assert detect_semester_change("IІ семестр") == 2
    assert detect_semester_change("2 семестр") == 2
    assert detect_semester_change("2-й семестр") == 2
    assert detect_semester_change("семестр ІІ") == 2

    assert detect_semester_change("Звичайна тема уроку") is None


def test_detect_folder_name_variants(tmp_path: Path) -> None:
    d1 = docx.Document()
    d1.add_paragraph("Календарно-тематичне планування для 8-9 класу")
    assert detect_folder_name(d1, tmp_path / "plan.docx") == "8-9 клас"

    d2 = docx.Document()
    d2.add_paragraph("Навчальна програма без згадки")
    assert detect_folder_name(d2, tmp_path / "Інформатика_7_клас.docx") == "7 клас"

    d3 = docx.Document()
    assert detect_folder_name(d3, tmp_path / "plan.docx") == DEFAULT_FOLDER_NAME


def test_decode_dropped_path_handles_utf8_and_cp1251(tmp_path: Path) -> None:
    cyr_file = tmp_path / "План_уроків.docx"
    cyr_file.write_bytes(b"stub")

    raw_str = str(cyr_file)
    assert decode_dropped_path(raw_str) == raw_str
    assert decode_dropped_path(raw_str.encode("utf-8")) == raw_str
    assert decode_dropped_path(raw_str.encode("cp1251")) == raw_str


def test_extract_planning_from_docx_ignores_non_lesson_numbers(tmp_path: Path) -> None:
    src = _create_sample_docx(tmp_path / "ktp_8.docx", title="Планування для 8 класу")
    plan = extract_planning_from_docx(src)

    assert plan.folder_name == "8 клас"
    assert plan.total_count == 3
    assert plan.semester1_count == 2
    assert plan.semester2_count == 1

    assert plan.lessons[0] == Lesson(
        semester=1,
        module_num=1,
        module_name="Тема 1. Основи алгоритмізації",
        lesson_num=1,
        text="Поняття алгоритму та виконавця",
    )
    assert plan.lessons[1].lesson_num == 2
    assert plan.lessons[2].semester == 2
    assert plan.lessons[2].module_num == 2
    assert plan.lessons[2].module_name == "Варіативний модуль: Python"


def test_create_excel_for_semester_clones_styles_over_100_rows(tmp_path: Path) -> None:
    lessons = [
        Lesson(
            semester=1,
            module_num=1,
            module_name="Великий модуль",
            lesson_num=idx,
            text=f"Урок №{idx}",
        )
        for idx in range(1, 116)
    ]
    out_xlsx = tmp_path / "large_sem1.xlsx"
    written = create_excel_for_semester(
        lessons,
        out_xlsx,
        target_sem=1,
        template_path=resolve_template_path(),
    )

    assert written == 115
    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["теми"]
    assert ws.cell(6, 8).value == "Урок №1"
    assert ws.cell(120, 1).value == 115
    assert ws.cell(120, 7).value == 115
    assert ws.cell(120, 8).value == "Урок №115"
    assert ws.cell(120, 8).font.name == ws.cell(105, 8).font.name


def test_convert_docx_to_excel_end_to_end(tmp_path: Path) -> None:
    src = _create_sample_docx(tmp_path / "plan_test.docx", title="Планування для 9 класу")
    result = convert_docx_to_excel(
        src,
        semester_choice="both",
        output_parent_dir=tmp_path,
        template_path=resolve_template_path(),
    )

    assert Path(result.output_dir) == tmp_path / "9 клас"
    assert len(result.created_files) == 2
    assert result.lessons_exported == 3

    for f in result.created_files:
        assert Path(f).is_file()


def test_bundled_examples_are_consistent() -> None:
    repo_root = Path(__file__).resolve().parent.parent
    example_docx = repo_root / "examples" / "sample_ktp_8_klas.docx"
    example_s1 = repo_root / "examples" / "sample_ktp_8_klas_1_sem.xlsx"
    example_s2 = repo_root / "examples" / "sample_ktp_8_klas_2_sem.xlsx"

    if not example_docx.is_file():
        pytest.skip("Example files not generated yet")

    plan = extract_planning_from_docx(example_docx)
    assert plan.folder_name == "8 клас"
    assert plan.semester1_count > 0
    assert plan.semester2_count > 0

    wb1 = openpyxl.load_workbook(example_s1)
    wb2 = openpyxl.load_workbook(example_s2)
    assert wb1["теми"].cell(6, 8).value == plan.lessons[0].text
    assert wb2["теми"].cell(6, 2).value == 2
