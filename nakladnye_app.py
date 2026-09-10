# -*- coding: utf-8 -*-
"""
Оконное приложение для обработки накладных (Витебские ковры + Обои УПД).

Одна кнопка «Обработать файлы» запускает обработку всех новых накладных
в выбранной папке, ход работы построчно выводится в журнал ниже, а по
окончании показывается сводка о готовности.

Запуск из исходников:   python nakladnye_app.py
Сборка в один файл:     см. build_exe.bat (Windows) / build.sh (Linux)
"""

import datetime
import os
import queue
import subprocess
import sys
import threading
import traceback

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

import convert_nakladnaya as core

APP_NAME = "Накладные: Витебск + Обои"
APP_VERSION = "1.0"

BG = "#f4f5f7"
CARD = "#ffffff"
BORDER = "#d8dbe0"
TEXT = "#1f2430"
MUTED = "#6b7280"
ACCENT = "#1f6feb"
ACCENT_HOVER = "#1a5fd0"
ACCENT_OFF = "#9db8e8"

LOG_BG = "#111721"
LOG_FG = "#dfe4ec"

STATUS_COLORS = {
    "idle": ("#eceef2", "#4b5563"),
    "run": ("#e8f0fe", "#1a4fa0"),
    "ok": ("#e7f6ec", "#1b6b3a"),
    "warn": ("#fdf3e2", "#8a5a10"),
    "err": ("#fdeceb", "#a3251c"),
}


def open_in_explorer(path):
    """Открыть папку в системном файловом менеджере."""
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def pick_font(candidates, size, **kw):
    available = set(tkfont.families())
    for name in candidates:
        if name in available:
            return tkfont.Font(family=name, size=size, **kw)
    return tkfont.Font(size=size, **kw)


def plural(n, one, few, many):
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


class QueueWriter:
    """Подменяет stdout/stderr, чтобы посторонний вывод тоже попадал в журнал."""

    def __init__(self, put, level="info"):
        self._put = put
        self._level = level
        self._buf = ""

    def write(self, text):
        self._buf += text
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self._put(line, self._level)

    def flush(self):
        if self._buf:
            self._put(self._buf, self._level)
            self._buf = ""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} — v{APP_VERSION}")
        self.geometry("940x680")
        self.minsize(760, 560)
        self.configure(bg=BG)

        self.messages = queue.Queue()
        self.worker = None
        self.folder = tk.StringVar(value=core.app_dir())
        self.force = tk.BooleanVar(value=False)

        self.f_ui = pick_font(["Segoe UI", "Inter", "DejaVu Sans", "Helvetica"], 10)
        self.f_ui_bold = pick_font(["Segoe UI", "Inter", "DejaVu Sans", "Helvetica"], 10, weight="bold")
        self.f_title = pick_font(["Segoe UI Semibold", "Segoe UI", "Inter", "DejaVu Sans"], 15, weight="bold")
        self.f_btn = pick_font(["Segoe UI Semibold", "Segoe UI", "Inter", "DejaVu Sans"], 11, weight="bold")
        self.f_log = pick_font(["Consolas", "Cascadia Mono", "DejaVu Sans Mono", "Courier New"], 10)
        self.f_log_bold = pick_font(["Consolas", "Cascadia Mono", "DejaVu Sans Mono", "Courier New"], 10, weight="bold")
        self.f_mono_small = pick_font(["Consolas", "DejaVu Sans Mono", "Courier New"], 9)

        self._init_theme()
        self._build_ui()

        sys.stdout = QueueWriter(self._put)
        sys.stderr = QueueWriter(self._put, "err")
        core.set_log_sink(self._put)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(80, self._drain)
        self._greeting()

    # ------------------------------------------------------------------ тема
    def _init_theme(self):
        style = ttk.Style(self)
        for theme in ("vista", "clam", "default"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("TLabel", background=BG, foreground=TEXT, font=self.f_ui)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=self.f_ui)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=self.f_ui)
        style.configure("CardMuted.TLabel", background=CARD, foreground=MUTED, font=self.f_ui)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=self.f_title)
        style.configure("Path.TLabel", background=CARD, foreground=TEXT, font=self.f_mono_small)
        style.configure("TCheckbutton", background=BG, foreground=TEXT, font=self.f_ui)
        style.configure("TButton", font=self.f_ui, padding=(12, 6))
        style.configure("Thin.Horizontal.TProgressbar", thickness=6,
                        background=ACCENT, troughcolor=BORDER, borderwidth=0)

    # -------------------------------------------------------------- интерфейс
    def _build_ui(self):
        root = ttk.Frame(self, padding=(18, 16, 18, 14))
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(5, weight=1)

        head = ttk.Frame(root)
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(0, weight=1)
        ttk.Label(head, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(head, text="Накладные поставщиков → файл для загрузки в учётную программу",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 0))

        # --- папка с накладными
        folder_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER,
                               highlightthickness=1, bd=0)
        folder_card.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        folder_card.columnconfigure(1, weight=1)
        ttk.Label(folder_card, text="Папка с накладными:",
                  style="CardMuted.TLabel").grid(row=0, column=0, padx=(12, 8), pady=10)
        ttk.Label(folder_card, textvariable=self.folder, style="Path.TLabel",
                  anchor="w").grid(row=0, column=1, sticky="ew", pady=10)
        ttk.Button(folder_card, text="Выбрать…", command=self._choose_folder,
                   width=12).grid(row=0, column=2, padx=(8, 10), pady=8)

        # --- кнопка запуска
        actions = ttk.Frame(root)
        actions.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        actions.columnconfigure(2, weight=1)

        self.run_btn = tk.Button(
            actions, text="Обработать файлы", command=self._start,
            font=self.f_btn, bg=ACCENT, fg="white",
            activebackground=ACCENT_HOVER, activeforeground="white",
            disabledforeground="#eef2fb", relief="flat", bd=0,
            padx=28, pady=12, cursor="hand2",
        )
        self.run_btn.grid(row=0, column=0, sticky="w")
        self.run_btn.bind("<Enter>", lambda _e: self._hover(True))
        self.run_btn.bind("<Leave>", lambda _e: self._hover(False))

        ttk.Checkbutton(actions, text="Обрабатывать повторно (не учитывать журнал)",
                        variable=self.force).grid(row=0, column=1, padx=(16, 0))

        self.progress = ttk.Progressbar(root, mode="indeterminate",
                                        style="Thin.Horizontal.TProgressbar")
        self.progress.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self.progress.grid_remove()

        # --- строка состояния
        self.status_box = tk.Frame(root, bg=STATUS_COLORS["idle"][0])
        self.status_box.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self.status_lbl = tk.Label(self.status_box, text="", bg=STATUS_COLORS["idle"][0],
                                   fg=STATUS_COLORS["idle"][1], font=self.f_ui_bold,
                                   anchor="w", padx=12, pady=9, justify="left")
        self.status_lbl.pack(fill="x")
        self._set_status("Готов к работе. Нажмите «Обработать файлы».", "idle")

        # --- журнал
        log_head = ttk.Frame(root)
        log_head.grid(row=4, column=0, sticky="ew", pady=(14, 6))
        log_head.columnconfigure(0, weight=1)
        ttk.Label(log_head, text="Журнал обработки",
                  style="TLabel").grid(row=0, column=0, sticky="w")

        log_wrap = tk.Frame(root, bg=BORDER, bd=0, highlightthickness=0)
        log_wrap.grid(row=5, column=0, sticky="nsew")
        log_wrap.columnconfigure(0, weight=1)
        log_wrap.rowconfigure(0, weight=1)

        self.log = tk.Text(log_wrap, bg=LOG_BG, fg=LOG_FG, font=self.f_log,
                           insertbackground=LOG_FG, relief="flat", bd=0,
                           padx=12, pady=10, wrap="word", height=14,
                           state="disabled", selectbackground="#2c3a52")
        self.log.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        bar = ttk.Scrollbar(log_wrap, orient="vertical", command=self.log.yview)
        bar.grid(row=0, column=1, sticky="ns", padx=(0, 1), pady=1)
        self.log.configure(yscrollcommand=bar.set)

        self.log.tag_configure("info", foreground=LOG_FG)
        self.log.tag_configure("muted", foreground="#8b95a7")
        self.log.tag_configure("head", foreground="#8ab4ff", font=self.f_log_bold)
        self.log.tag_configure("ok", foreground="#7ee2a8")
        self.log.tag_configure("warn", foreground="#f5c86b")
        self.log.tag_configure("err", foreground="#ff8f85", font=self.f_log_bold)
        self.log.tag_configure("rule", foreground="#3a465c")

        # --- нижние кнопки
        bottom = ttk.Frame(root)
        bottom.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        bottom.columnconfigure(3, weight=1)
        self.open_btn = ttk.Button(bottom, text="Открыть папку с результатами",
                                   command=self._open_folder)
        self.open_btn.grid(row=0, column=0)
        ttk.Button(bottom, text="Сохранить журнал…",
                   command=self._save_log).grid(row=0, column=1, padx=8)
        ttk.Button(bottom, text="Очистить",
                   command=self._clear_log).grid(row=0, column=2)
        ttk.Label(bottom, text=f"v{APP_VERSION}", style="Muted.TLabel").grid(row=0, column=3, sticky="e")

    def _hover(self, on):
        if str(self.run_btn["state"]) != "disabled":
            self.run_btn.configure(bg=ACCENT_HOVER if on else ACCENT)

    # ------------------------------------------------------------------ журнал
    def _put(self, text, level="info"):
        """Потокобезопасно: вызывается и из рабочего потока."""
        self.messages.put((text, level))

    def _drain_now(self):
        try:
            while True:
                text, level = self.messages.get_nowait()
                self._append(text, level)
        except queue.Empty:
            pass

    def _drain(self):
        self._drain_now()
        self.after(80, self._drain)

    def _append(self, text, level="info"):
        tag = level if level in ("info", "muted", "head", "ok", "warn", "err", "rule") else "info"
        if set(text.strip()) in ({"-"}, {"="}, {"─"}) and text.strip():
            tag = "rule"
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.configure(state="disabled")
        self.log.see("end")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _greeting(self):
        self._append(f"{APP_NAME}, версия {APP_VERSION}", "head")
        self._append("Приложение обрабатывает два вида накладных:", "muted")
        self._append("  • Витебские ковры — фактура .xls", "muted")
        self._append("  • Обои — сырой УПД .xlsx (со словом «Обои» в наименованиях)", "muted")
        self._append("", "muted")
        self._append("Положите накладные в папку, указанную выше, и нажмите", "muted")
        self._append("«Обработать файлы». Результаты появятся в той же папке", "muted")
        self._append("с суффиксом «_Обработано».", "muted")
        self._append("=" * 60)

    def _set_status(self, text, kind="idle"):
        bg, fg = STATUS_COLORS.get(kind, STATUS_COLORS["idle"])
        self.status_box.configure(bg=bg)
        self.status_lbl.configure(text=text, bg=bg, fg=fg)

    # ------------------------------------------------------------------ папка
    def _choose_folder(self):
        if self.worker and self.worker.is_alive():
            return
        chosen = filedialog.askdirectory(title="Папка с накладными",
                                         initialdir=self.folder.get())
        if chosen:
            self.folder.set(os.path.abspath(chosen))
            self._append(f"Рабочая папка: {self.folder.get()}", "muted")

    def _open_folder(self):
        path = self.folder.get()
        if not os.path.isdir(path):
            messagebox.showwarning(APP_NAME, f"Папка не найдена:\n{path}")
            return
        try:
            open_in_explorer(path)
        except Exception as exc:
            messagebox.showwarning(APP_NAME, f"Не удалось открыть папку:\n{exc}")

    def _save_log(self):
        default = f"журнал_{datetime.datetime.now():%Y-%m-%d_%H-%M}.txt"
        path = filedialog.asksaveasfilename(
            title="Сохранить журнал", initialfile=default,
            initialdir=self.folder.get(), defaultextension=".txt",
            filetypes=[("Текстовый файл", "*.txt"), ("Все файлы", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.log.get("1.0", "end"))
        except Exception as exc:
            messagebox.showwarning(APP_NAME, f"Не удалось сохранить журнал:\n{exc}")
            return
        self._append(f"Журнал сохранён: {path}", "ok")

    # ------------------------------------------------------------------ работа
    def _start(self):
        if self.worker and self.worker.is_alive():
            return
        folder = self.folder.get()
        if not os.path.isdir(folder):
            self._set_status(f"Папка не найдена: {folder}", "err")
            messagebox.showwarning(APP_NAME, f"Папка не найдена:\n{folder}")
            return

        self.run_btn.configure(state="disabled", bg=ACCENT_OFF, text="Обработка…")
        self.status_box.grid_remove()
        self.progress.grid()
        self.progress.start(12)
        self._set_status("Обработка запущена…", "run")

        self._append("", "info")
        self._append(f"Запуск обработки — {datetime.datetime.now():%d.%m.%Y %H:%M:%S}", "head")
        self._append("=" * 60)

        force = self.force.get()
        self.worker = threading.Thread(target=self._work, args=(folder, force), daemon=True)
        self.worker.start()

    def _work(self, folder, force):
        try:
            itog = core.process_folder(folder, force=force)
        except Exception:
            self._put("НЕОЖИДАННАЯ ОШИБКА:", "err")
            for line in traceback.format_exc().rstrip().splitlines():
                self._put("  " + line, "err")
            itog = None
        self.after(0, self._finish, itog)

    def _finish(self, itog):
        # сводка должна идти после последних строк, ещё лежащих в очереди
        self._drain_now()
        self.progress.stop()
        self.progress.grid_remove()
        self.status_box.grid()
        self.run_btn.configure(state="normal", bg=ACCENT, text="Обработать файлы")

        if itog is None:
            self._append("=" * 60)
            self._append("ОБРАБОТКА ПРЕРВАНА ОШИБКОЙ", "err")
            self._set_status("Обработка прервана ошибкой. Подробности в журнале.", "err")
            return

        self._report(itog)

    def _report(self, itog):
        done = itog["done"]
        errors = itog["errors"]
        skipped = itog["skipped"]

        self._append("=" * 60)
        if errors:
            self._append(f"ЗАВЕРШЕНО С ОШИБКАМИ — {datetime.datetime.now():%H:%M:%S}", "err")
        elif done:
            self._append(f"ГОТОВО — {datetime.datetime.now():%H:%M:%S}", "ok")
        else:
            self._append(f"НЕЧЕГО ОБРАБАТЫВАТЬ — {datetime.datetime.now():%H:%M:%S}", "warn")

        self._append(f"Обработано файлов: {len(done)}")
        self._append(f"Перенесено строк: {itog['rows']}")

        if done:
            self._append("Созданные файлы:")
            for item in done:
                rows = item["rows"]
                word = plural(rows, "строка", "строки", "строк")
                self._append(f"   • {os.path.basename(item['out'])} — {rows} {word}", "ok")
        if skipped:
            self._append(f"Пропущено (обработаны ранее): {len(skipped)}", "warn")
        if errors:
            self._append(f"Ошибок: {len(errors)}", "err")
            for name, msg in errors:
                self._append(f"   • {name}: {msg}", "err")

        self._append(f"Время работы: {itog['seconds']:.1f} с".replace(".", ","))
        self._append(f"Папка результатов: {itog['folder']}", "muted")
        self._append("=" * 60)
        self._append("")

        text = core.opisanie_itoga(itog)
        if errors:
            kind = "err"
            text += ". Подробности в журнале."
        elif done:
            kind = "ok"
            files = len(done)
            text = (f"Готово. Обработано {files} {plural(files, 'файл', 'файла', 'файлов')}, "
                    f"перенесено {itog['rows']} {plural(itog['rows'], 'строка', 'строки', 'строк')}. "
                    f"Результаты — в папке с накладными.")
        else:
            kind = "warn"
        self._set_status(text, kind)

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askokcancel(APP_NAME, "Обработка ещё идёт. Закрыть приложение?"):
                return
        self.destroy()


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
