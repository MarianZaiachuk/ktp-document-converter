#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Скрипт автоматичної збірки Конвертер_КТП.exe для Windows."""

from __future__ import annotations

import importlib.util
import os
import shutil
import subprocess
import sys

REQUIRED_MODULES = {
    "PyInstaller": "pyinstaller",
    "docx": "python-docx",
    "openpyxl": "openpyxl",
}
if sys.platform == "win32":
    REQUIRED_MODULES["windnd"] = "windnd"


def ensure_dependencies(force_install: bool = False) -> None:
    """Install missing build/runtime dependencies via pip only when needed or requested."""
    missing = [
        pkg for mod, pkg in REQUIRED_MODULES.items() if importlib.util.find_spec(mod) is None
    ]
    if not force_install and not missing:
        print("Крок 1/3: Усі необхідні бібліотеки вже встановлено.")
        return

    packages = list(REQUIRED_MODULES.values()) if force_install else missing
    print(f"Крок 1/3: Встановлення бібліотек ({', '.join(packages)})...")
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "--upgrade", *packages],
        check=True,
    )


def main() -> int:
    print("=" * 65)
    print("   ЗБІРКА АВТОНОМНОЇ ПРОГРАМИ: Конвертер_КТП.exe (v3.1)")
    print("=" * 65)
    print()

    force_install = "--install-deps" in sys.argv[1:]
    try:
        ensure_dependencies(force_install=force_install)
    except subprocess.CalledProcessError as exc:
        print(f"Помилка встановлення бібліотек: {exc}")
        return 1

    template_src = os.path.join("templates", "topics-template.xlsx")
    if not os.path.isfile(template_src):
        print(f"Помилка: не знайдено шаблон '{template_src}'.")
        return 1

    print()
    print("Крок 2/3: Збірка програми в один .exe файл без консолі...")
    add_data_arg = f"{template_src}{os.pathsep}templates"
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconsole",
        "--clean",
        "--add-data",
        add_data_arg,
        "--name=Конвертер_КТП",
        "gui_app.py",
    ]

    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as exc:
        print(f"Помилка під час збірки PyInstaller: {exc}")
        return 1

    print()
    print("Крок 3/3: Копіювання готового файлу...")
    exe_src = os.path.join("dist", "Конвертер_КТП.exe")
    exe_dst = "Конвертер_КТП.exe"

    if os.path.exists(exe_src):
        shutil.copy2(exe_src, exe_dst)
        print("=" * 65)
        print("🎉 УСПІХ! Програму успішно створено:")
        print(f"   --> {os.path.abspath(exe_dst)}")
        print()
        print("Тепер ви можете скопіювати файл 'Конвертер_КТП.exe' на флешку чи надіслати вчителям.")
        print("Вчителям НЕ потрібен Python - програма запускається подвійним кліком!")
        print("=" * 65)
    else:
        print("Попередження: перевірте папку dist/")

    print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
