"""Graphical interface for Excel Legacy Converter."""

import os
import queue
import threading
import tkinter as tk
import tkinter.scrolledtext as scrolledtext
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from excel_converter import __version__
from excel_converter.cli import run_conversion
from excel_converter.discovery import build_tasks, find_files


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Excel Legacy Converter  v{__version__}")
        self.resizable(True, True)
        self.minsize(560, 520)

        try:
            ttk.Style().theme_use("vista")
        except tk.TclError:
            pass

        self._input_mode = tk.StringVar(value="files")
        self._recursive_var = tk.BooleanVar(value=False)
        self._no_excel_var = tk.BooleanVar(value=False)

        # Internally stored selections (Paths)
        self._selected_files: list[Path] = []   # used in "files" mode
        self._selected_folder: Path | None = None  # used in "folder" mode
        self._output_folder: Path | None = None    # None = use default

        self._queue: queue.Queue = queue.Queue()
        self._stop_event = threading.Event()

        self._build_ui()

    # ── Build UI ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}
        self._build_input_section(self).pack(fill=tk.X, **pad)
        self._build_output_section(self).pack(fill=tk.X, **pad)
        self._build_action_section(self).pack(fill=tk.X, **pad)
        self._build_progress_section(self).pack(fill=tk.BOTH, expand=True, **pad)

    def _build_input_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Archivos de origen")
        inner = ttk.Frame(frame)
        inner.pack(fill=tk.X, padx=8, pady=6)

        # Mode radio buttons
        mode_row = ttk.Frame(inner)
        mode_row.pack(fill=tk.X)
        self._radio_files = ttk.Radiobutton(
            mode_row, text="Archivos individuales", variable=self._input_mode,
            value="files", command=self._on_mode_change,
        )
        self._radio_files.pack(side=tk.LEFT)
        self._radio_folder = ttk.Radiobutton(
            mode_row, text="Carpeta completa", variable=self._input_mode,
            value="folder", command=self._on_mode_change,
        )
        self._radio_folder.pack(side=tk.LEFT, padx=(16, 0))

        # Selection row: button + path label
        sel_row = ttk.Frame(inner)
        sel_row.pack(fill=tk.X, pady=(6, 0))
        self._btn_browse_input = ttk.Button(
            sel_row, text="Seleccionar archivos...", width=22,
            command=self._on_browse_input,
        )
        self._btn_browse_input.pack(side=tk.LEFT)
        self._btn_clear_input = ttk.Button(
            sel_row, text="✕", width=3, command=self._on_clear_input,
        )
        self._btn_clear_input.pack(side=tk.LEFT, padx=(4, 0))
        self._input_label = ttk.Label(
            sel_row, text="Sin selección", foreground="#888888",
            font=("Segoe UI", 9), anchor=tk.W,
        )
        self._input_label.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)

        # Recursive checkbox (only active in folder mode)
        self._recursive_chk = ttk.Checkbutton(
            inner, text="Incluir subcarpetas", variable=self._recursive_var,
            state=tk.DISABLED,
        )
        self._recursive_chk.pack(anchor=tk.W, pady=(6, 0))

        return frame

    def _build_output_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Carpeta de destino")
        inner = ttk.Frame(frame)
        inner.pack(fill=tk.X, padx=8, pady=6)

        row = ttk.Frame(inner)
        row.pack(fill=tk.X)
        self._btn_browse_output = ttk.Button(
            row, text="Seleccionar destino...", width=22,
            command=self._on_browse_output,
        )
        self._btn_browse_output.pack(side=tk.LEFT)
        self._btn_reset_output = ttk.Button(
            row, text="↺", width=3, command=self._on_reset_output,
        )
        self._btn_reset_output.pack(side=tk.LEFT, padx=(4, 0))
        self._output_label = ttk.Label(
            row, text=self._default_output_text(), foreground="#888888",
            font=("Segoe UI", 9), anchor=tk.W,
        )
        self._output_label.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)

        return frame

    def _build_action_section(self, parent):
        frame = ttk.Frame(parent)

        self._no_excel_chk = ttk.Checkbutton(
            frame, text="Modo sin Excel  (solo datos, sin formato de celdas)",
            variable=self._no_excel_var,
        )
        self._no_excel_chk.pack(anchor=tk.W, pady=(0, 6))

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)
        self._btn_convert = ttk.Button(
            btn_row, text="Convertir", width=16, command=self._on_convert,
        )
        self._btn_convert.pack(side=tk.LEFT)
        self._btn_cancel = ttk.Button(
            btn_row, text="Cancelar", width=10,
            command=self._on_cancel, state=tk.DISABLED,
        )
        self._btn_cancel.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(btn_row, text=f"v{__version__}", foreground="#aaaaaa").pack(side=tk.RIGHT)

        return frame

    def _build_progress_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Progreso")
        inner = ttk.Frame(frame)
        inner.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        # Progress bar + counter
        pb_row = ttk.Frame(inner)
        pb_row.pack(fill=tk.X)
        self._progress_var = tk.DoubleVar(value=0.0)
        self._progressbar = ttk.Progressbar(
            pb_row, variable=self._progress_var,
            maximum=100.0, mode="determinate",
        )
        self._progressbar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._progress_label = ttk.Label(pb_row, text="0 / 0", width=9, anchor=tk.E)
        self._progress_label.pack(side=tk.LEFT, padx=(6, 0))

        # Log area
        self._log_text = scrolledtext.ScrolledText(
            inner, height=11, state=tk.DISABLED,
            font=("Consolas", 9), bg="#1a1a2e", fg="#d0d0d0",
            relief=tk.FLAT, padx=6, pady=4, insertwidth=0,
        )
        self._log_text.tag_configure("ok",   foreground="#66bb6a")
        self._log_text.tag_configure("fail", foreground="#ef5350")
        self._log_text.tag_configure("info", foreground="#78909c")
        self._log_text.tag_configure("warn", foreground="#ffa726")
        self._log_text.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        # Status line
        self._status_label = ttk.Label(inner, text="Listo.", foreground="#666666")
        self._status_label.pack(anchor=tk.W, pady=(4, 0))

        return frame

    # ── Event handlers ─────────────────────────────────────────────────────────

    def _on_mode_change(self):
        self._selected_files.clear()
        self._selected_folder = None
        self._input_label.config(text="Sin selección", foreground="#888888")
        if self._input_mode.get() == "files":
            self._btn_browse_input.config(text="Seleccionar archivos...")
            self._recursive_chk.config(state=tk.DISABLED)
        else:
            self._btn_browse_input.config(text="Seleccionar carpeta...")
            self._recursive_chk.config(state=tk.NORMAL)
        self._refresh_output_label()

    def _on_browse_input(self):
        if self._input_mode.get() == "files":
            paths = filedialog.askopenfilenames(
                title="Seleccionar archivos .xls",
                filetypes=[("Excel 97-2003", "*.xls"), ("Todos los archivos", "*.*")],
            )
            if not paths:
                return
            self._selected_files = [Path(p) for p in paths]
            n = len(self._selected_files)
            label = (
                str(self._selected_files[0]) if n == 1
                else f"{n} archivos seleccionados"
            )
            self._input_label.config(text=label, foreground="#cccccc")
        else:
            folder = filedialog.askdirectory(title="Seleccionar carpeta de origen")
            if not folder:
                return
            self._selected_folder = Path(folder)
            self._input_label.config(text=str(self._selected_folder), foreground="#cccccc")
        self._refresh_output_label()

    def _on_clear_input(self):
        self._selected_files.clear()
        self._selected_folder = None
        self._input_label.config(text="Sin selección", foreground="#888888")
        self._refresh_output_label()

    def _on_browse_output(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if not folder:
            return
        self._output_folder = Path(folder)
        self._output_label.config(text=str(self._output_folder), foreground="#cccccc")

    def _on_reset_output(self):
        self._output_folder = None
        self._refresh_output_label()

    def _on_convert(self):
        mode = self._input_mode.get()

        # Build task list
        if mode == "files":
            if not self._selected_files:
                messagebox.showwarning("Sin selección", "Selecciona al menos un archivo .xls.")
                return
            files = [f for f in self._selected_files if f.is_file()]
            if not files:
                messagebox.showerror("Error", "Ninguno de los archivos seleccionados existe.")
                return
            output_dir = self._get_output_dir(files[0].parent)
            tasks = [(src, output_dir / src.name) for src in files]
        else:
            if self._selected_folder is None:
                messagebox.showwarning("Sin selección", "Selecciona una carpeta de origen.")
                return
            if not self._selected_folder.is_dir():
                messagebox.showerror("Error", f"La carpeta ya no existe:\n{self._selected_folder}")
                return
            recursive = self._recursive_var.get()
            xls_files = find_files(self._selected_folder, recursive)
            if not xls_files:
                messagebox.showinfo(
                    "Sin archivos",
                    "No se encontraron archivos .xls en la carpeta seleccionada.",
                )
                return
            output_dir = self._get_output_dir(self._selected_folder)
            tasks = build_tasks(xls_files, self._selected_folder, output_dir)

        # Reset progress UI
        self._log_clear()
        self._progress_var.set(0.0)
        self._progress_label.config(text=f"0 / {len(tasks)}")
        self._status_label.config(text="Iniciando conversión...")
        self._set_busy(True)

        self._queue = queue.Queue()
        self._stop_event = threading.Event()

        threading.Thread(
            target=self._run_in_thread,
            args=(tasks, self._no_excel_var.get(), os.cpu_count() or 4),
            daemon=True,
        ).start()
        self.after(100, self._poll_queue)

    def _on_cancel(self):
        self._stop_event.set()
        self._status_label.config(text="Cancelando...")
        self._btn_cancel.config(state=tk.DISABLED)

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _get_output_dir(self, fallback_base: Path) -> Path:
        return self._output_folder if self._output_folder else fallback_base / "converted"

    def _default_output_text(self) -> str:
        return "Por defecto: <origen>/converted/"

    def _refresh_output_label(self):
        if self._output_folder:
            self._output_label.config(text=str(self._output_folder), foreground="#cccccc")
        else:
            self._output_label.config(text=self._default_output_text(), foreground="#888888")

    def _set_busy(self, busy: bool):
        state_all = tk.DISABLED if busy else tk.NORMAL
        for w in (self._btn_convert, self._btn_browse_input, self._btn_clear_input,
                  self._btn_browse_output, self._btn_reset_output,
                  self._no_excel_chk, self._radio_files, self._radio_folder):
            w.config(state=state_all)
        # Restore mode-specific recursive checkbox state when un-busying
        if not busy:
            rec_state = tk.NORMAL if self._input_mode.get() == "folder" else tk.DISABLED
            self._recursive_chk.config(state=rec_state)
        else:
            self._recursive_chk.config(state=tk.DISABLED)
        self._btn_cancel.config(state=tk.NORMAL if busy else tk.DISABLED)

    def _log_clear(self):
        self._log_text.config(state=tk.NORMAL)
        self._log_text.delete("1.0", tk.END)
        self._log_text.config(state=tk.DISABLED)

    def _log(self, text: str, tag: str | None = None):
        self._log_text.config(state=tk.NORMAL)
        if tag:
            self._log_text.insert(tk.END, text, tag)
        else:
            self._log_text.insert(tk.END, text)
        self._log_text.see(tk.END)
        self._log_text.config(state=tk.DISABLED)

    # ── Background thread ──────────────────────────────────────────────────────

    def _run_in_thread(self, tasks, no_excel: bool, workers: int):
        try:
            ok = 0
            failed = 0
            gen = run_conversion(tasks, no_excel=no_excel, workers=workers)
            start = next(gen)  # consume mode-info event
            self._queue.put({"type": "start", **start})
            for p in gen:
                if self._stop_event.is_set():
                    break
                if p["success"]:
                    ok += 1
                else:
                    failed += 1
                self._queue.put({"type": "progress", **p})
            self._queue.put({"type": "done", "ok": ok, "failed": failed})
        except Exception as exc:
            self._queue.put({"type": "error", "exc": exc})

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                t = msg.get("type")
                if t == "start":
                    self._handle_start(msg)
                elif t == "progress":
                    self._handle_progress(msg)
                elif t == "done":
                    self._handle_done(msg)
                    return
                elif t == "error":
                    self._handle_error(msg["exc"])
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_start(self, msg: dict):
        if msg["com_unavailable"]:
            self._log("Nota: Excel COM no disponible — usando modo sin formato.\n", "warn")
        mode_text = "Excel COM (formato completo)" if msg["mode"] == "com" else "Sin Excel (solo datos)"
        self._log(f"Modo: {mode_text}\n", "info")
        self._log(f"Total: {msg['total']} archivo(s)\n\n", "info")
        self._status_label.config(text="Convirtiendo...")

    def _handle_progress(self, msg: dict):
        src: Path = msg["src"]
        done: int = msg["done"]
        total: int = msg["total"]

        pct = done / total * 100 if total else 0
        self._progress_var.set(pct)
        self._progress_label.config(text=f"{done} / {total}")
        self._status_label.config(text=f"{src.name}")

        if msg["success"]:
            fmt_tag = f"  [{msg['fmt'].upper()}]" if msg["fmt"] in ("html", "xml") else ""
            self._log(f"  [OK]   {src.name}{fmt_tag}\n", "ok")
        else:
            self._log(f"  [FAIL] {src.name}: {msg['error']}\n", "fail")

    def _handle_done(self, msg: dict):
        ok: int = msg["ok"]
        failed: int = msg["failed"]
        self._set_busy(False)
        self._progress_var.set(100.0)
        summary = f"\nListo: {ok} convertido(s), {failed} fallido(s)."
        self._log(summary + "\n", "fail" if failed else "ok")
        self._status_label.config(text=summary.strip())

    def _handle_error(self, exc: Exception):
        self._set_busy(False)
        self._log(f"\nError inesperado: {exc}\n", "fail")
        self._status_label.config(text="Error inesperado.")
        messagebox.showerror("Error", f"Error inesperado durante la conversión:\n{exc}")


def main():
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
