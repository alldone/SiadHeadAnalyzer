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

from .telemedicina_core import (
    generate_trace,
    load_vertical_config,
    package_resource,
    validate_source_workbook,
)


STATE_FILE = Path.home() / ".siad_head_analyzer_telemedicina.json"


class TelemedicinaGuiApp:
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

        if not embed_mode:
            self.root.title("Telemedicina - Tracciato PNT")
            self.root.geometry("1220x820")
            self.root.minsize(1040, 700)

        self.source_var = self.tk.StringVar(value="")
        self.output_var = self.tk.StringVar(value="")
        self.config_var = self.tk.StringVar(value=str(package_resource("config.json")))
        self.status_var = self.tk.StringVar(
            value="Seleziona il file Excel sorgente. Le colonne saranno individuate per intestazione."
        )
        self._busy = False
        self._last_output: Path | None = None

        self._build_ui()
        self._load_state()

    def _build_ui(self) -> None:
        if not self.embed_mode:
            self._build_menu()
        frame = self.ttk.Frame(self.parent, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ribbon = self.ttk.Frame(frame)
        ribbon.grid(row=0, column=0, sticky="ew")
        for column in range(3):
            ribbon.columnconfigure(column, weight=1)
        self._ribbon_group(
            ribbon,
            0,
            "File",
            [("Sorgente Excel", self.choose_source), ("File output", self.choose_output)],
        )
        self._ribbon_group(
            ribbon,
            1,
            "Controlli",
            [("Verifica sorgente", self.validate_source), ("Genera tracciato", self.run_generation)],
        )
        self._ribbon_group(
            ribbon,
            2,
            "Risultati",
            [("Apri Excel", self.open_output), ("Apri cartella", self.open_output_dir)],
        )

        paths = self.ttk.LabelFrame(frame, text="Configurazione", padding=12)
        paths.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        paths.columnconfigure(1, weight=1)
        self._path_row(paths, 0, "File Excel sorgente", self.source_var, self.choose_source)
        self._path_row(paths, 1, "File Excel output", self.output_var, self.choose_output)
        self._path_row(paths, 2, "Traduzioni strutture", self.config_var, self.choose_config)

        hint = (
            "Il verticale legge per nome soltanto: ID paziente, regione di residenza, ASL di residenza, "
            "fascia d'età, sesso ed erogante/erogatore. Le altre colonne sono ignorate. "
            "F contiene l'azienda di assistenza; P e S sono ricondotte all'azienda e alla struttura erogante."
        )
        self.ttk.Label(frame, text=hint, wraplength=1160, justify="left").grid(
            row=2, column=0, sticky="ew", pady=(0, 8)
        )

        self.notebook = self.ttk.Notebook(frame)
        self.notebook.grid(row=3, column=0, sticky="nsew")
        self.output_text = self._text_tab("Output")
        self.validation_text = self._text_tab("Validazione sorgente")

        status = self.ttk.Frame(frame)
        status.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        status.columnconfigure(1, weight=1)
        self.progress = self.ttk.Progressbar(status, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.ttk.Label(status, textvariable=self.status_var, anchor="w").grid(
            row=0, column=1, sticky="ew"
        )

    def _build_menu(self) -> None:
        menu = self.tk.Menu(self.root)
        file_menu = self.tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Seleziona sorgente Excel", command=self.choose_source)
        file_menu.add_command(label="Seleziona output Excel", command=self.choose_output)
        file_menu.add_command(label="Seleziona configurazione", command=self.choose_config)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.destroy)
        menu.add_cascade(label="File", menu=file_menu)
        run_menu = self.tk.Menu(menu, tearoff=False)
        run_menu.add_command(label="Verifica sorgente", command=self.validate_source)
        run_menu.add_command(label="Genera tracciato", command=self.run_generation)
        menu.add_cascade(label="Esecuzione", menu=run_menu)
        self.root.config(menu=menu)

    def _ribbon_group(
        self, parent: Any, column: int, title: str, buttons: list[tuple[str, Any]]
    ) -> None:
        group = self.ttk.LabelFrame(parent, text=title, padding=8)
        group.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column < 2 else 0)
        for index, (label, command) in enumerate(buttons):
            self.ttk.Button(group, text=label, command=command).grid(
                row=0, column=index, padx=(0, 8), pady=2
            )

    def _path_row(self, parent: Any, row: int, label: str, variable: Any, command: Any) -> None:
        self.ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        self.ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", padx=8, pady=4
        )
        self.ttk.Button(parent, text="Sfoglia...", command=command).grid(
            row=row, column=2, sticky="e", pady=4
        )

    def _text_tab(self, title: str) -> Any:
        tab = self.ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab, text=title)
        widget = self.scrolledtext.ScrolledText(tab, wrap="word", font=("Menlo", 11))
        widget.pack(fill="both", expand=True)
        widget.configure(state="disabled")
        widget.tag_configure("error", foreground="#9c0006")
        widget.tag_configure("success", foreground="#0b6b2f")
        widget.tag_configure("header", foreground="#1f4e78", font=("Menlo", 11, "bold"))
        return widget

    def choose_source(self) -> None:
        current = self.source_var.get().strip()
        path = self.filedialog.askopenfilename(
            initialdir=str(Path(current).parent) if current else str(Path.home()),
            filetypes=[("File Excel", "*.xlsx *.xlsm"), ("Tutti i file", "*.*")],
        )
        if not path:
            return
        self.source_var.set(path)
        source = Path(path)
        self.output_var.set(str(source.with_name(f"Tracciato_PNT_{source.stem}.xlsx")))
        self._save_state()
        self.validate_source(show_dialog=False)

    def choose_output(self) -> None:
        current = self.output_var.get().strip()
        path = self.filedialog.asksaveasfilename(
            initialdir=str(Path(current).parent) if current else str(Path.home()),
            initialfile=Path(current).name if current else "Tracciato_PNT.xlsx",
            defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)
            self._save_state()

    def choose_config(self) -> None:
        current = self.config_var.get().strip()
        path = self.filedialog.askopenfilename(
            initialdir=str(Path(current).parent) if current else str(Path.home()),
            filetypes=[("Configurazione JSON", "*.json"), ("Tutti i file", "*.*")],
        )
        if path:
            self.config_var.set(path)
            self._save_state()

    def _configuration(self) -> dict[str, Any]:
        path = self.config_var.get().strip()
        if not path:
            raise RuntimeError("Selezionare il file JSON con le traduzioni delle strutture.")
        return load_vertical_config(path)

    def validate_source(self, *, show_dialog: bool = True) -> bool:
        source = self.source_var.get().strip()
        if not source:
            if show_dialog:
                self.messagebox.showerror("Sorgente mancante", "Seleziona il file Excel sorgente.")
            return False
        try:
            cfg = self._configuration()
            validation = validate_source_workbook(source, cfg)
            message = validation.message(cfg["source_columns"])
        except Exception as exc:
            validation = None
            message = f"Configurazione o sorgente non valida:\n{exc}"
        self._set_text(self.validation_text, message)
        self.notebook.select(1)
        if validation and validation.valid:
            self.status_var.set("Sorgente conforme: tutte le colonne utilizzate sono presenti.")
            if show_dialog:
                self.messagebox.showinfo("Verifica sorgente", "File conforme alle colonne richieste.")
            return True
        self.status_var.set("Sorgente non conforme: consultare la scheda Validazione sorgente.")
        if show_dialog:
            self.messagebox.showerror("Sorgente non conforme", message)
        return False

    def run_generation(self) -> None:
        if self._busy or not self.validate_source(show_dialog=True):
            return
        source = self.source_var.get().strip()
        output = self.output_var.get().strip()
        if not output:
            self.messagebox.showerror("Output mancante", "Indica il file Excel di output.")
            return
        output_path = Path(output)
        if output_path.exists() and not self.messagebox.askyesno(
            "Sovrascrivere output?", f"Il file esiste già:\n{output_path}\n\nSovrascriverlo?"
        ):
            return
        try:
            cfg = self._configuration()
        except Exception as exc:
            self.messagebox.showerror("Configurazione non valida", str(exc))
            return

        self._save_state()
        self._set_text(self.output_text, "")
        self.notebook.select(0)
        self._set_busy(True, "Generazione del tracciato in corso...")
        threading.Thread(target=self._worker, args=(source, output, cfg), daemon=True).start()

    def _worker(self, source: str, output: str, cfg: dict[str, Any]) -> None:
        try:
            summary = generate_trace(
                source,
                output,
                cfg=cfg,
                log=lambda message: self._append(message, "header" if message.startswith("Sorgente") else None),
            )
            self._last_output = summary.output_path
            report = (
                "\nGenerazione completata.\n"
                f"Casi sorgente: {summary.total_cases}\n"
                f"Righe aggregate: {summary.aggregated_rows}\n"
                f"Totale colonna H: {summary.output_quantity_total}\n"
                f"File: {summary.output_path}"
            )
            self._append(report, "success")
            self.root.after(0, lambda: self.status_var.set(
                f"Completato: totale colonna H = {summary.output_quantity_total}."
            ))
            self.root.after(0, lambda: self.messagebox.showinfo(
                "Tracciato generato",
                f"Tracciato generato correttamente.\nTotale colonna H: {summary.output_quantity_total}",
            ))
        except Exception as exc:
            detail = str(exc)
            self._append("\n" + traceback.format_exc(), "error")
            self.root.after(0, lambda: self._show_generation_error(detail))
        finally:
            self.root.after(0, lambda: self._set_busy(False))

    def _show_generation_error(self, detail: str) -> None:
        self._set_text(self.validation_text, detail)
        self.notebook.select(1)
        self.status_var.set("Elaborazione sospesa: correggere gli elementi indicati.")
        self.messagebox.showerror("Elaborazione sospesa", detail)

    def open_output(self) -> None:
        raw_path = self.output_var.get().strip()
        if self._last_output is None and not raw_path:
            self.messagebox.showerror("Output mancante", "Nessun file di output selezionato.")
            return
        path = self._last_output or Path(raw_path)
        self._open_path(path)

    def open_output_dir(self) -> None:
        raw_path = self.output_var.get().strip()
        if self._last_output is None and not raw_path:
            self.messagebox.showerror("Output mancante", "Nessun file di output selezionato.")
            return
        path = self._last_output or Path(raw_path)
        self._open_path(path.parent)

    def _open_path(self, path: Path) -> None:
        if not path.exists():
            self.messagebox.showerror("Percorso non trovato", str(path))
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        elif sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", str(path)], check=False)

    def _append(self, message: str, tag: str | None = None) -> None:
        self.root.after(0, lambda: self._append_now(message + "\n", tag))

    def _append_now(self, message: str, tag: str | None) -> None:
        self.output_text.configure(state="normal")
        self.output_text.insert("end", message, tag or ())
        self.output_text.see("end")
        self.output_text.configure(state="disabled")

    @staticmethod
    def _set_text(widget: Any, text: str) -> None:
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
        else:
            self.progress.stop()

    def _save_state(self) -> None:
        try:
            STATE_FILE.write_text(
                json.dumps(
                    {
                        "version": self.STATE_VERSION,
                        "source": self.source_var.get().strip(),
                        "output": self.output_var.get().strip(),
                        "config": self.config_var.get().strip(),
                    },
                    indent=2,
                ),
                encoding="utf-8",
            )
        except OSError:
            pass

    def _load_state(self) -> None:
        try:
            payload = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for key, variable in (
            ("source", self.source_var),
            ("output", self.output_var),
            ("config", self.config_var),
        ):
            value = payload.get(key)
            if value and (key != "config" or Path(value).is_file()):
                variable.set(value)


def main() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    root = tk.Tk()
    TelemedicinaGuiApp(root, tk, ttk, filedialog, messagebox, scrolledtext)
    root.mainloop()


if __name__ == "__main__":
    main()
