#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графічна програма (GUI) для конвертації календарно-тематичного планування
із Word (.docx) у шаблон Excel (.xlsx) для системи «Єдина школа».
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import (
    ParsedPlan,
    convert_docx_to_excel,
    decode_dropped_path,
    extract_planning_from_docx,
)

# --- High-DPI чіткість у Windows ---
if sys.platform == "win32":
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# --- Підтримка Drag & Drop через windnd (якщо встановлено) ---
try:
    import windnd

    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False


class KTPConverterApp(tk.Tk):
    """Desktop Tkinter GUI for batch KTP (.docx -> .xlsx) conversion."""

    def __init__(self) -> None:
        super().__init__()
        self._scale = max(1.0, self.winfo_fpixels("1i") / 96.0)
        self.title("Конвертер КТП для «Єдиної школи» (v3.1)")
        self.geometry(f"{self._px(820)}x{self._px(660)}")
        self.minsize(self._px(760), self._px(580))
        self.configure(bg="#f1f5f9")

        self.selected_files: list[str] = []
        self._normalized_paths: set[str] = set()
        self.parsed_plans: dict[str, ParsedPlan] = {}
        self.last_result_dirs: list[str] = []
        self._is_converting = False

        self.apply_theme()
        self.setup_ui()
        self.setup_drag_and_drop()

    def _px(self, value: int) -> int:
        return round(value * self._scale)

    def apply_theme(self) -> None:
        self.style = ttk.Style(self)
        available = self.style.theme_names()
        if "vista" in available:
            self.style.theme_use("vista")
        elif "clam" in available:
            self.style.theme_use("clam")
        self.style.configure("Treeview", rowheight=self._px(26), font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        self.style.configure("TButton", padding=(self._px(10), self._px(4)))

    def setup_ui(self) -> None:
        # 1. Header Banner
        header = tk.Frame(self, bg="#1e3a8a", height=self._px(88))
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title = tk.Label(
            header,
            text="📚 Конвертер календарно-тематичного планування",
            font=("Segoe UI", 16, "bold"),
            fg="#ffffff",
            bg="#1e3a8a",
        )
        title.pack(pady=(self._px(12), self._px(2)))

        sub = tk.Label(
            header,
            text="Автоматичне перенесення КТП з Word у шаблон Excel системи «Єдина школа»",
            font=("Segoe UI", 10),
            fg="#93c5fd",
            bg="#1e3a8a",
        )
        sub.pack()

        # 2. Main content container
        container = tk.Frame(self, bg="#f1f5f9", padx=self._px(20), pady=self._px(12))
        container.pack(fill=tk.BOTH, expand=True)

        # Card 1: File selection & List (Drag & Drop Zone)
        card1 = tk.LabelFrame(
            container,
            text=" 1. Оберіть або перетягніть файли планування (.docx) ",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff",
            fg="#1e293b",
            padx=self._px(12),
            pady=self._px(10),
            relief=tk.GROOVE,
        )
        card1.pack(fill=tk.BOTH, expand=True, pady=(0, self._px(10)))

        btn_bar = tk.Frame(card1, bg="#ffffff")
        btn_bar.pack(fill=tk.X, pady=(0, self._px(4)))

        add_btn = ttk.Button(btn_bar, text="📂 Додати файл(и)...", command=self.browse_files)
        add_btn.pack(side=tk.LEFT, padx=(0, self._px(8)))

        clear_btn = ttk.Button(btn_bar, text="🗑️ Очистити список", command=self.clear_files)
        clear_btn.pack(side=tk.LEFT)

        dnd_hint = tk.Label(
            card1,
            text="💡 Можна виділити кілька файлів одразу або перетягнути їх сюди мишкою",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff",
            anchor=tk.W,
            justify=tk.LEFT,
        )
        dnd_hint.pack(fill=tk.X, pady=(0, self._px(6)))

        tree_frame = tk.Frame(card1, bg="#ffffff")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("name", "class", "lessons", "path")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=5)
        self.tree.heading("name", text="Файл планування (.docx)")
        self.tree.heading("class", text="Розпізнаний клас")
        self.tree.heading("lessons", text="Знайдено уроків")
        self.tree.heading("path", text="Розташування")

        self.tree.column("name", width=self._px(230), minwidth=self._px(160), anchor=tk.W)
        self.tree.column("class", width=self._px(145), minwidth=self._px(125), anchor=tk.CENTER)
        self.tree.column("lessons", width=self._px(155), minwidth=self._px(130), anchor=tk.CENTER)
        self.tree.column("path", width=self._px(210), minwidth=self._px(140), anchor=tk.W)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.files_count_lbl = tk.Label(
            card1,
            text="Файлів не додано. Натисніть «Додати файл(и)» або перетягніть файли Word у це вікно.",
            font=("Segoe UI", 9),
            fg="#64748b",
            bg="#ffffff",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=self._px(720),
        )
        self.files_count_lbl.pack(fill=tk.X, pady=(self._px(6), 0))

        # Card 2: Options
        card2 = tk.LabelFrame(
            container,
            text=" 2. Налаштування формування ",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff",
            fg="#1e293b",
            padx=self._px(12),
            pady=self._px(8),
            relief=tk.GROOVE,
        )
        card2.pack(fill=tk.X, pady=(0, self._px(10)))

        self.sem_var = tk.StringVar(value="both")
        r1 = tk.Radiobutton(
            card2,
            text="Сформувати обидва семестри окремими файлами (рекомендовано)",
            variable=self.sem_var,
            value="both",
            font=("Segoe UI", 10),
            bg="#ffffff",
            activebackground="#ffffff",
        )
        r2 = tk.Radiobutton(
            card2,
            text="Тільки I семестр",
            variable=self.sem_var,
            value="1",
            font=("Segoe UI", 10),
            bg="#ffffff",
            activebackground="#ffffff",
        )
        r3 = tk.Radiobutton(
            card2,
            text="Тільки II семестр",
            variable=self.sem_var,
            value="2",
            font=("Segoe UI", 10),
            bg="#ffffff",
            activebackground="#ffffff",
        )

        r1.pack(anchor=tk.W, pady=1)
        r2.pack(anchor=tk.W, pady=1)
        r3.pack(anchor=tk.W, pady=1)

        folder_info_lbl = tk.Label(
            card2,
            text="📁 Для кожного плану автоматично створиться папка класу (наприклад, «8-9 клас» або «Календарне планування»), а всередині будуть файли Excel.",
            font=("Segoe UI", 9),
            fg="#0369a1",
            bg="#ffffff",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=self._px(720),
        )
        folder_info_lbl.pack(fill=tk.X, pady=(self._px(4), 0))

        # 3. Action button
        self.convert_btn = tk.Button(
            container,
            text="⚡ Конвертувати всі додані плани",
            font=("Segoe UI", 12, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            cursor="hand2",
            relief=tk.FLAT,
            command=self.run_conversion,
        )
        self.convert_btn.pack(fill=tk.X, ipady=self._px(9))

        # 4. Status and Open directory button
        self.status_lbl = tk.Label(
            container,
            text="",
            font=("Segoe UI", 10),
            fg="#334155",
            bg="#f1f5f9",
            wraplength=self._px(720),
            justify=tk.CENTER,
        )
        self.status_lbl.pack(pady=(self._px(6), self._px(2)))

        self.open_dir_btn = ttk.Button(
            container,
            text="📁 Відкрити папку зі створеними файлами",
            command=self.open_output_dir,
        )
        self.open_dir_btn.pack()
        self.open_dir_btn.pack_forget()

        card1.bind(
            "<Configure>",
            lambda e: self.files_count_lbl.config(
                wraplength=max(self._px(300), e.width - self._px(32))
            ),
        )
        card2.bind(
            "<Configure>",
            lambda e: folder_info_lbl.config(
                wraplength=max(self._px(300), e.width - self._px(32))
            ),
        )
        container.bind(
            "<Configure>",
            lambda e: self.status_lbl.config(
                wraplength=max(self._px(300), e.width - self._px(32))
            ),
        )

    def setup_drag_and_drop(self) -> None:
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self, func=self.on_drop_files)
            except Exception:
                pass

    def on_drop_files(self, files: list[str | bytes]) -> None:
        new_docs: list[str] = []
        for raw in files:
            decoded = decode_dropped_path(raw)
            if decoded.lower().endswith(".docx"):
                new_docs.append(decoded)
        if new_docs:
            self.add_files_to_list(new_docs)

    def browse_files(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="Оберіть файл(и) календарного планування Word",
            filetypes=[("Документи Word", "*.docx"), ("Усі файли", "*.*")],
        )
        if filenames:
            self.add_files_to_list(list(filenames))

    def add_files_to_list(self, file_paths: list[str]) -> None:
        for raw_path in file_paths:
            abs_path = os.path.abspath(raw_path)
            norm_path = os.path.normcase(abs_path)
            if norm_path in self._normalized_paths:
                continue

            self._normalized_paths.add(norm_path)
            self.selected_files.append(abs_path)

            try:
                plan = extract_planning_from_docx(abs_path)
                self.parsed_plans[norm_path] = plan
                folder_name = plan.folder_name
                if plan.total_count > 0:
                    les_info = f"{plan.total_count} ур. ({plan.semester1_count} + {plan.semester2_count})"
                else:
                    les_info = "0 ур. (таблицю не знайдено)"
            except Exception:
                folder_name = "Невідомо"
                les_info = "Помилка файлу"

            self.tree.insert(
                "",
                tk.END,
                values=(os.path.basename(abs_path), folder_name, les_info, os.path.dirname(abs_path)),
            )

        cnt = len(self.selected_files)
        self.files_count_lbl.config(
            text=f"✅ Додано планів: {cnt}. Готово до конвертації.",
            fg="#16a34a",
        )
        self.open_dir_btn.pack_forget()

    def clear_files(self) -> None:
        if self._is_converting:
            return
        self.selected_files.clear()
        self._normalized_paths.clear()
        self.parsed_plans.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.files_count_lbl.config(
            text="Файлів не додано. Натисніть «Додати файл(и)» або перетягніть файли Word у це вікно.",
            fg="#64748b",
        )
        self.status_lbl.config(text="")
        self.open_dir_btn.pack_forget()

    def run_conversion(self) -> None:
        if self._is_converting:
            return
        if not self.selected_files:
            messagebox.showwarning("Увага", "Будь ласка, додайте хоча б один файл Word (.docx)!")
            return

        choice = self.sem_var.get()
        files_snapshot = list(self.selected_files)
        plans_snapshot = dict(self.parsed_plans)

        self._is_converting = True
        self.convert_btn.config(state=tk.DISABLED, text="⏳ Триває конвертація...")
        self.status_lbl.config(text="⏳ Триває обробка планів та створення папок...", fg="#2563eb")
        self.open_dir_btn.pack_forget()

        worker = threading.Thread(
            target=self._convert_worker,
            args=(files_snapshot, plans_snapshot, choice),
            daemon=True,
        )
        worker.start()

    def _convert_worker(
        self,
        file_paths: list[str],
        plans_cache: dict[str, ParsedPlan],
        choice: str,
    ) -> None:
        result_dirs: list[str] = []
        success_count = 0
        total_created_files = 0
        summary_lines: list[str] = []

        for doc_path in file_paths:
            try:
                norm_path = os.path.normcase(os.path.abspath(doc_path))
                cached_plan = plans_cache.get(norm_path)
                result = convert_docx_to_excel(
                    doc_path,
                    semester_choice=choice,
                    parsed_plan=cached_plan,
                )

                if not result.created_files:
                    summary_lines.append(f"❌ {os.path.basename(doc_path)}: уроків не знайдено.")
                    continue

                result_dirs.append(result.output_dir)
                success_count += 1
                created_count = len(result.created_files)
                total_created_files += created_count
                folder_label = os.path.basename(result.output_dir)
                summary_lines.append(
                    f"✅ {os.path.basename(doc_path)} ➔ папка «{folder_label}» ({created_count} файл.)"
                )
            except Exception as exc:
                summary_lines.append(f"❌ {os.path.basename(doc_path)}: {exc}")

        self.after(
            0,
            lambda: self._on_conversion_complete(
                total_inputs=len(file_paths),
                result_dirs=result_dirs,
                success_count=success_count,
                total_created_files=total_created_files,
                summary_lines=summary_lines,
            ),
        )

    def _on_conversion_complete(
        self,
        total_inputs: int,
        result_dirs: list[str],
        success_count: int,
        total_created_files: int,
        summary_lines: list[str],
    ) -> None:
        self._is_converting = False
        self.convert_btn.config(state=tk.NORMAL, text="⚡ Конвертувати всі додані плани")
        self.last_result_dirs = result_dirs

        if success_count > 0:
            self.status_lbl.config(
                text=(
                    f"🎉 Успішно оброблено {success_count} з {total_inputs} планів "
                    f"(створено {total_created_files} Excel-файлів у відповідних папках класів)."
                ),
                fg="#16a34a",
            )
            self.open_dir_btn.pack(pady=(4, 0))
            messagebox.showinfo(
                "Успіх!",
                "Конвертацію успішно завершено!\n\n" + "\n".join(summary_lines),
            )
        else:
            self.status_lbl.config(text="Не вдалося конвертувати файли.", fg="#dc2626")
            messagebox.showerror(
                "Помилка",
                "Не вдалося конвертувати жодного файлу.\n\n" + "\n".join(summary_lines),
            )

    def open_output_dir(self) -> None:
        if self.last_result_dirs:
            target = self.last_result_dirs[0]
            if os.path.exists(target):
                if sys.platform == "win32":
                    os.startfile(target)
                elif sys.platform == "darwin":
                    subprocess.run(["open", target], check=False)
                else:
                    subprocess.run(["xdg-open", target], check=False)


if __name__ == "__main__":
    app = KTPConverterApp()
    app.mainloop()
