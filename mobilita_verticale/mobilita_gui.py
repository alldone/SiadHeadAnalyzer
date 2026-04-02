#!/usr/bin/env python3

from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import traceback
from typing import Any

try:
    from .mobilita_report import (
        COLUMN_DEFINITIONS,
        DEFAULT_SELECTED_COLUMN_KEYS,
        ensure_available_output_path,
        generate_report_from_directories,
        next_output_path,
        validate_input_directories,
    )
except ImportError:
    from mobilita_report import (
        COLUMN_DEFINITIONS,
        DEFAULT_SELECTED_COLUMN_KEYS,
        ensure_available_output_path,
        generate_report_from_directories,
        next_output_path,
        validate_input_directories,
    )


PROJECT_ROOT = Path(__file__).resolve().parent
STATE_FILE = PROJECT_ROOT / ".mobilita_gui_state.json"


class MobilitaGuiApp:
    SPINNER_FRAMES = ["[...]", "[ ..]", "[  .]", "[ ..]"]
    STATE_VERSION = 1

    def __init__(
        self,
        root: Any,
        tk_module: Any,
        ttk_module: Any,
        filedialog_module: Any,
        messagebox_module: Any,
        scrolledtext_module: Any,
        *,
        parent: Any | None = None,
        embed_mode: bool = False,
    ) -> None:
        self.root = root
        self.parent = parent or root
        self.embed_mode = embed_mode
        self.tk = tk_module
        self.ttk = ttk_module
        self.filedialog = filedialog_module
        self.messagebox = messagebox_module
        self.scrolledtext = scrolledtext_module

        if not self.embed_mode:
            self.root.title("Mobilita Infraregionale Farmaci")
            self.root.geometry("1220x820")
            self.root.minsize(1080, 720)

        self.attiva_dir_var = self.tk.StringVar(value="")
        self.passiva_dir_var = self.tk.StringVar(value="")
        self.output_file_var = self.tk.StringVar(value="")
        self.status_var = self.tk.StringVar(
            value="Seleziona le cartelle attiva e passiva, poi genera il report Excel."
        )
        self.spinner_var = self.tk.StringVar(value="")
        self.column_options = list(COLUMN_DEFINITIONS)

        self._busy = False
        self._spinner_index = 0
        self._spinner_job: str | None = None
        self._last_output: str | None = None

        self._build_ui()
        self._load_state()

    def _build_ui(self) -> None:
        if not self.embed_mode:
            self._build_menu()

        style = self.ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        frame = self.ttk.Frame(self.parent, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ribbon = self.ttk.Frame(frame)
        ribbon.grid(row=0, column=0, sticky="ew")
        ribbon.columnconfigure(0, weight=1)
        ribbon.columnconfigure(1, weight=1)
        ribbon.columnconfigure(2, weight=1)

        self._build_ribbon_group(
            ribbon,
            0,
            "Percorsi",
            [
                ("Attiva", self.choose_attiva_dir),
                ("Passiva", self.choose_passiva_dir),
                ("File output", self.choose_output_file),
            ],
        )
        self._build_ribbon_group(
            ribbon,
            1,
            "Esecuzione",
            [("Genera report", self.run_report), ("Pulisci log", self.clear_views)],
        )
        self._build_ribbon_group(
            ribbon,
            2,
            "Report",
            [("Apri Excel", self.open_output_file), ("Apri cartella", self.open_output_dir)],
        )

        config_frame = self.ttk.LabelFrame(frame, text="Configurazione", padding=12)
        config_frame.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        config_frame.columnconfigure(1, weight=1)

        self._add_path_row(config_frame, 0, "Cartella attiva", self.attiva_dir_var, self.choose_attiva_dir)
        self._add_path_row(config_frame, 1, "Cartella passiva", self.passiva_dir_var, self.choose_passiva_dir)
        self._add_path_row(config_frame, 2, "File Excel output", self.output_file_var, self.choose_output_file)
        self._add_column_selector(config_frame, 3)

        hint = (
            "Seleziona separatamente le cartelle attiva e passiva. "
            "In ciascuna cartella il programma abbina i file '<azienda> importi.csv' e "
            "'<azienda> prestazioni.csv', riporta le tipologie selezionate nella listbox "
            "e genera un workbook con due fogli: attiva e passiva. "
            "La listbox delle tipologie e' multi-selezione; per default sono selezionate "
            "FARMACEUTICA e SOMM. DIRETTA DI FARMACI."
        )
        self.ttk.Label(frame, text=hint, wraplength=1160, justify="left").grid(row=2, column=0, sticky="ew", pady=(0, 8))

        self.notebook = self.ttk.Notebook(frame)
        self.notebook.grid(row=3, column=0, sticky="nsew")

        self.console_text = self._add_text_tab("Output")
        self.validation_text = self._add_text_tab("Validazione")

        status_bar = self.ttk.Frame(frame)
        status_bar.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        status_bar.columnconfigure(2, weight=1)

        self.ttk.Label(status_bar, textvariable=self.spinner_var, width=5, anchor="w").grid(row=0, column=0, sticky="w")
        self.progress = self.ttk.Progressbar(status_bar, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=1, sticky="w", padx=(0, 12))
        self.ttk.Label(status_bar, textvariable=self.status_var, anchor="w").grid(row=0, column=2, sticky="ew")

    def _build_menu(self) -> None:
        menu = self.tk.Menu(self.root)

        file_menu = self.tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Seleziona cartella attiva", command=self.choose_attiva_dir)
        file_menu.add_command(label="Seleziona cartella passiva", command=self.choose_passiva_dir)
        file_menu.add_command(label="Seleziona file output", command=self.choose_output_file)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        run_menu = self.tk.Menu(menu, tearoff=False)
        run_menu.add_command(label="Genera report", command=self.run_report)
        run_menu.add_command(label="Pulisci log", command=self.clear_views)
        menu.add_cascade(label="Esecuzione", menu=run_menu)

        report_menu = self.tk.Menu(menu, tearoff=False)
        report_menu.add_command(label="Apri file Excel", command=self.open_output_file)
        report_menu.add_command(label="Apri cartella output", command=self.open_output_dir)
        menu.add_cascade(label="Report", menu=report_menu)

        self.root.config(menu=menu)

    def _build_ribbon_group(self, parent: Any, column: int, title: str, buttons: list[tuple[str, Any]]) -> None:
        group = self.ttk.LabelFrame(parent, text=title, padding=8)
        group.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column < 2 else 0)
        for index, (label, command) in enumerate(buttons):
            self.ttk.Button(group, text=label, command=command).grid(row=0, column=index, padx=(0, 8), pady=2)

    def _add_path_row(self, parent: Any, row: int, label: str, variable: Any, command: Any) -> None:
        self.ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        self.ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        self.ttk.Button(parent, text="Sfoglia...", command=command).grid(row=row, column=2, sticky="e", pady=4)

    def _add_column_selector(self, parent: Any, row: int) -> None:
        self.ttk.Label(parent, text="Tipologie da riportare").grid(row=row, column=0, sticky="nw", pady=4)
        box_frame = self.ttk.Frame(parent)
        box_frame.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        box_frame.columnconfigure(0, weight=1)

        self.column_listbox = self.tk.Listbox(
            box_frame,
            selectmode="extended",
            exportselection=False,
            height=min(len(self.column_options), 8),
        )
        self.column_listbox.grid(row=0, column=0, sticky="ew")
        scrollbar = self.ttk.Scrollbar(box_frame, orient="vertical", command=self.column_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.column_listbox.configure(yscrollcommand=scrollbar.set)

        for definition in self.column_options:
            self.column_listbox.insert("end", definition.label)

        self._restore_column_selection(list(DEFAULT_SELECTED_COLUMN_KEYS))
        self.column_listbox.bind("<<ListboxSelect>>", lambda _event: self._save_state())

    def _add_text_tab(self, title: str) -> Any:
        tab = self.ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab, text=title)
        widget = self.scrolledtext.ScrolledText(tab, wrap="word", font=("Menlo", 11))
        widget.pack(fill="both", expand=True)
        widget.configure(state="disabled")
        widget.tag_configure("info", foreground="#1f1f1f")
        widget.tag_configure("warn", foreground="#9c6500")
        widget.tag_configure("error", foreground="#9c0006")
        widget.tag_configure("success", foreground="#0b6b2f")
        widget.tag_configure("header", foreground="#1f4e78", font=("Menlo", 11, "bold"))
        return widget

    def choose_attiva_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.attiva_dir_var.get() or str(Path.home()))
        if path:
            self.attiva_dir_var.set(path)
            self._suggest_output_file()
            self._save_state()

    def choose_passiva_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.passiva_dir_var.get() or str(Path.home()))
        if path:
            self.passiva_dir_var.set(path)
            self._suggest_output_file()
            self._save_state()

    def _suggest_output_file(self) -> None:
        if self.output_file_var.get().strip():
            return
        suggested_dir = self._preferred_output_dir()
        if suggested_dir is None:
            return
        self.output_file_var.set(str(next_output_path(suggested_dir)))

    def _preferred_output_dir(self) -> Path | None:
        attiva_value = self.attiva_dir_var.get().strip()
        passiva_value = self.passiva_dir_var.get().strip()

        attiva_path = Path(attiva_value).expanduser() if attiva_value else None
        passiva_path = Path(passiva_value).expanduser() if passiva_value else None

        if attiva_path and passiva_path and attiva_path.parent == passiva_path.parent:
            return attiva_path.parent
        if attiva_path:
            return attiva_path.parent
        if passiva_path:
            return passiva_path.parent
        return None

    def choose_output_file(self) -> None:
        current = self.output_file_var.get().strip()
        init_dir = str(Path(current).parent) if current else str(self._preferred_output_dir() or Path.home())
        path = self.filedialog.asksaveasfilename(
            initialdir=init_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.output_file_var.set(path)
            self._save_state()

    def _selected_column_keys(self) -> list[str]:
        return [self.column_options[index].key for index in self.column_listbox.curselection()]

    def _restore_column_selection(self, keys: list[str]) -> None:
        self.column_listbox.selection_clear(0, "end")
        selected_set = set(keys)
        for index, definition in enumerate(self.column_options):
            if definition.key in selected_set:
                self.column_listbox.selection_set(index)
        if not self.column_listbox.curselection():
            default_set = set(DEFAULT_SELECTED_COLUMN_KEYS)
            for index, definition in enumerate(self.column_options):
                if definition.key in default_set:
                    self.column_listbox.selection_set(index)

    def clear_views(self) -> None:
        for widget in (self.console_text, self.validation_text):
            self._set_text(widget, "")
        self.status_var.set("Log ripuliti.")

    def run_report(self) -> None:
        if self._busy:
            return

        attiva_dir = Path(self.attiva_dir_var.get().strip()).expanduser()
        passiva_dir = Path(self.passiva_dir_var.get().strip()).expanduser()
        output_file_value = self.output_file_var.get().strip()
        selected_column_keys = self._selected_column_keys()

        if not attiva_dir.exists() or not attiva_dir.is_dir():
            self.messagebox.showerror("Errore", "Seleziona una cartella attiva valida.")
            return
        if not passiva_dir.exists() or not passiva_dir.is_dir():
            self.messagebox.showerror("Errore", "Seleziona una cartella passiva valida.")
            return
        if not selected_column_keys:
            self.messagebox.showerror("Errore", "Seleziona almeno una tipologia da riportare nel report.")
            return
        if not output_file_value:
            self.messagebox.showerror("Errore", "Indica il percorso del file Excel di output.")
            return

        output_path = ensure_available_output_path(Path(output_file_value).expanduser())
        self.output_file_var.set(str(output_path))

        errors = validate_input_directories(attiva_dir, passiva_dir)
        if errors:
            self._set_text(self.validation_text, "ERRORI DI VALIDAZIONE INPUT:\n\n" + "\n".join(f"- {error}" for error in errors))
            self.notebook.select(1)
            self.messagebox.showerror("Errore", f"Le cartelle input presentano {len(errors)} errore/i.\nVedi tab Validazione.")
            return

        self._save_state()
        self.clear_views()
        self._set_busy(True, "Generazione report in corso...")

        self.append_console(f"Cartella attiva: {attiva_dir}", "header")
        self.append_console(f"Cartella passiva: {passiva_dir}", "header")
        self.append_console(
            "Tipologie: " + ", ".join(
                definition.label for definition in self.column_options if definition.key in selected_column_keys
            ),
            "header",
        )
        self.append_console(f"Output: {output_path}", "header")
        self.append_console("", "info")

        threading.Thread(
            target=self._run_worker,
            args=(attiva_dir, passiva_dir, output_path, selected_column_keys),
            daemon=True,
        ).start()

    def _run_worker(
        self,
        attiva_dir: Path,
        passiva_dir: Path,
        output_path: Path,
        selected_column_keys: list[str],
    ) -> None:
        try:
            generated = generate_report_from_directories(
                attiva_dir,
                passiva_dir,
                output_path,
                selected_column_keys=selected_column_keys,
                log=lambda message: self.append_console(message, "info"),
            )
            self._last_output = str(generated)
            self.append_console("", "info")
            self.append_console("Report generato con successo.", "success")
            self.root.after(0, lambda: self.status_var.set(f"Report generato: {generated}"))
            self.root.after(
                0,
                lambda: self._set_text(
                    self.validation_text,
                    "Validazione input superata.\n"
                    "Controllo importi input/output superato.\n\n"
                    "Workbook generato correttamente.\n"
                    f"Percorso: {generated}",
                ),
            )
        except Exception as exc:
            self.append_console(traceback.format_exc(), "error")
            self.root.after(0, lambda: self.status_var.set("Errore durante l'elaborazione."))
            self.root.after(
                0,
                lambda: self._set_text(
                    self.validation_text,
                    "Errore durante l'elaborazione.\n\n"
                    f"{exc}",
                ),
            )
        finally:
            self.root.after(0, lambda: self._set_busy(False))

    def open_output_file(self) -> None:
        path = self._last_output or self.output_file_var.get().strip()
        if path:
            self._open_path(Path(path))

    def open_output_dir(self) -> None:
        path = self._last_output or self.output_file_var.get().strip()
        if path:
            self._open_path(Path(path).parent)

    def _open_path(self, path: Path) -> None:
        if not path.exists():
            self.messagebox.showerror("Errore", f"Percorso non trovato:\n{path}")
            return
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            elif sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:
            self.messagebox.showerror("Errore", f"Impossibile aprire il percorso:\n{exc}")

    def append_console(self, message: str, tag: str = "info") -> None:
        self.root.after(0, lambda: self._append_text(self.console_text, message + "\n", tag))

    def _append_text(self, widget: Any, text: str, tag: str) -> None:
        widget.configure(state="normal")
        widget.insert("end", text, tag)
        widget.see("end")
        widget.configure(state="disabled")

    def _set_text(self, widget: Any, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def _set_busy(self, busy: bool, status: str | None = None) -> None:
        self._busy = busy
        if status:
            self.status_var.set(status)
        if busy:
            self.progress.start(12)
            self._start_spinner()
        else:
            self.progress.stop()
            self._stop_spinner()

    def _start_spinner(self) -> None:
        if self._spinner_job is not None:
            return
        self._tick_spinner()

    def _tick_spinner(self) -> None:
        self.spinner_var.set(self.SPINNER_FRAMES[self._spinner_index % len(self.SPINNER_FRAMES)])
        self._spinner_index += 1
        self._spinner_job = self.root.after(160, self._tick_spinner)

    def _stop_spinner(self) -> None:
        if self._spinner_job is not None:
            self.root.after_cancel(self._spinner_job)
            self._spinner_job = None
        self.spinner_var.set("")

    def _save_state(self) -> None:
        payload = {
            "version": self.STATE_VERSION,
            "attiva_dir": self.attiva_dir_var.get().strip(),
            "passiva_dir": self.passiva_dir_var.get().strip(),
            "output_file": self.output_file_var.get().strip(),
            "selected_column_keys": self._selected_column_keys(),
        }
        STATE_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _load_state(self) -> None:
        if not STATE_FILE.exists():
            return
        try:
            payload = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        attiva_dir = payload.get("attiva_dir")
        passiva_dir = payload.get("passiva_dir")
        output_file = payload.get("output_file")
        selected_column_keys = payload.get("selected_column_keys")
        if attiva_dir:
            self.attiva_dir_var.set(attiva_dir)
        if passiva_dir:
            self.passiva_dir_var.set(passiva_dir)
        if output_file:
            self.output_file_var.set(output_file)
        if isinstance(selected_column_keys, list):
            self._restore_column_selection([str(key) for key in selected_column_keys])
        else:
            self._restore_column_selection(list(DEFAULT_SELECTED_COLUMN_KEYS))


def main() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    except ModuleNotFoundError as exc:
        raise SystemExit(f"Tkinter non disponibile: {exc}")

    root = tk.Tk()
    MobilitaGuiApp(root, tk, ttk, filedialog, messagebox, scrolledtext)
    root.mainloop()


if __name__ == "__main__":
    main()
