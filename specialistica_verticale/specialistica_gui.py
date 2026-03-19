#!/usr/bin/env python3

from __future__ import annotations

import csv
import json
import os
from collections import Counter
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
import subprocess
import sys
import threading
import traceback
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parent
STATE_FILE = PROJECT_ROOT / ".specialistica_gui_state.json"
DEFAULT_INPUT_DIR = PROJECT_ROOT / "bancadati"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"
DEFAULT_MAPPING_FILE = Path(os.environ.get("SPECIALISTICA_MAPPING_FILE", "")).expanduser() if os.environ.get("SPECIALISTICA_MAPPING_FILE", "").strip() else PROJECT_ROOT / "BRANCA_Codici regionali-Codici SSN.xlsx"
DEFAULT_INPUT_VALUE = str(DEFAULT_INPUT_DIR) if DEFAULT_INPUT_DIR.exists() else ""
DEFAULT_OUTPUT_VALUE = str(DEFAULT_OUTPUT_DIR) if DEFAULT_OUTPUT_DIR.exists() else ""
DEFAULT_MAPPING_VALUE = str(DEFAULT_MAPPING_FILE) if DEFAULT_MAPPING_FILE.exists() else ""


class TkLogStream:
    def __init__(self, app: "SpecialisticaGuiApp", tag: str) -> None:
        self.app = app
        self.tag = tag
        self._buffer = ""

    def write(self, text: str) -> int:
        if not text:
            return 0
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self.app.append_console(line, self.tag)
        return len(text)

    def flush(self) -> None:
        if self._buffer.strip():
            self.app.append_console(self._buffer.strip(), self.tag)
        self._buffer = ""


class SpecialisticaGuiApp:
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
            self.root.title("ETL Specialistica")
            self.root.geometry("1220x820")
            self.root.minsize(1080, 720)

        self.input_dir_var = self.tk.StringVar(value=DEFAULT_INPUT_VALUE)
        self.output_dir_var = self.tk.StringVar(value=DEFAULT_OUTPUT_VALUE)
        self.mapping_file_var = self.tk.StringVar(value=DEFAULT_MAPPING_VALUE)
        self.status_var = self.tk.StringVar(
            value="Seleziona input, output e, se necessario, la banca dati di trascodifica; poi avvia ETL + validazione."
        )
        self.spinner_var = self.tk.StringVar(value="")

        self._busy = False
        self._spinner_index = 0
        self._spinner_job: str | None = None

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
                ("Input", self.choose_input_dir),
                ("Output", self.choose_output_dir),
                ("Trascodifica", self.choose_mapping_file),
            ],
        )
        self._build_ribbon_group(
            ribbon,
            1,
            "Esecuzione",
            [
                ("ETL + Validazione", self.run_pipeline),
                ("Pulisci log", self.clear_views),
            ],
        )
        self._build_ribbon_group(
            ribbon,
            2,
            "Report",
            [
                ("Apri output", self.open_output_dir),
                ("Apri anomalie", self.open_anomaly_excel),
                ("Apri validazione", self.open_validation_summary),
            ],
        )

        path_frame = self.ttk.LabelFrame(frame, text="Configurazione", padding=12)
        path_frame.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        path_frame.columnconfigure(1, weight=1)

        self._add_path_row(path_frame, 0, "Cartella input", self.input_dir_var, self.choose_input_dir)
        self._add_path_row(path_frame, 1, "Cartella output", self.output_dir_var, self.choose_output_dir)
        self._add_path_row(path_frame, 2, "Banca dati trascodifica", self.mapping_file_var, self.choose_mapping_file, file_mode=True)

        hint = (
            "Il flusso esegue sempre prima la preparazione dei file e poi la validazione. "
            "La banca dati di trascodifica puo essere indicata esplicitamente; se lasciata vuota, viene usata "
            "quella definita da SPECIALISTICA_MAPPING_FILE oppure il file incluso nel verticale. "
            "I report delle anomalie e della validazione vengono riletti e mostrati nelle tab sottostanti."
        )
        self.ttk.Label(frame, text=hint, wraplength=1160, justify="left").grid(row=2, column=0, sticky="ew", pady=(0, 8))

        self.notebook = self.ttk.Notebook(frame)
        self.notebook.grid(row=3, column=0, sticky="nsew")

        self.console_text = self._add_text_tab("Output")
        self.anomalies_text = self._add_text_tab("Anomalie")
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
        file_menu.add_command(label="Seleziona input", command=self.choose_input_dir)
        file_menu.add_command(label="Seleziona output", command=self.choose_output_dir)
        file_menu.add_command(label="Seleziona trascodifica", command=self.choose_mapping_file)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        run_menu = self.tk.Menu(menu, tearoff=False)
        run_menu.add_command(label="ETL + Validazione", command=self.run_pipeline)
        run_menu.add_command(label="Pulisci log", command=self.clear_views)
        menu.add_cascade(label="Esecuzione", menu=run_menu)

        report_menu = self.tk.Menu(menu, tearoff=False)
        report_menu.add_command(label="Apri cartella output", command=self.open_output_dir)
        report_menu.add_command(label="Apri report anomalie", command=self.open_anomaly_excel)
        report_menu.add_command(label="Apri riepilogo validazione", command=self.open_validation_summary)
        menu.add_cascade(label="Report", menu=report_menu)

        self.root.config(menu=menu)

    def _build_ribbon_group(self, parent: Any, column: int, title: str, buttons: list[tuple[str, Any]]) -> None:
        group = self.ttk.LabelFrame(parent, text=title, padding=8)
        group.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column < 2 else 0)
        for index, (label, command) in enumerate(buttons):
            self.ttk.Button(group, text=label, command=command).grid(row=0, column=index, padx=(0, 8), pady=2)

    def _add_path_row(self, parent: Any, row: int, label: str, variable: Any, command: Any, file_mode: bool = False) -> None:
        self.ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        self.ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        self.ttk.Button(parent, text="Sfoglia...", command=command).grid(row=row, column=2, sticky="e", pady=4)
        if file_mode:
            self.ttk.Button(parent, text="Pulisci", command=lambda: variable.set("")).grid(row=row, column=3, sticky="e", padx=(8, 0), pady=4)

    def _add_text_tab(self, title: str):
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

    def choose_input_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.input_dir_var.get() or str(PROJECT_ROOT))
        if path:
            self.input_dir_var.set(path)
            self._save_state()

    def choose_output_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.output_dir_var.get() or str(PROJECT_ROOT))
        if path:
            self.output_dir_var.set(path)
            self._save_state()

    def choose_mapping_file(self) -> None:
        path = self.filedialog.askopenfilename(
            initialdir=str(Path(self.mapping_file_var.get()).expanduser().parent) if self.mapping_file_var.get().strip() else str(PROJECT_ROOT),
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.mapping_file_var.set(path)
            self._save_state()

    def clear_views(self) -> None:
        for widget in (self.console_text, self.anomalies_text, self.validation_text):
            self._set_text(widget, "")
        self.status_var.set("Log ripuliti.")

    def run_pipeline(self) -> None:
        if self._busy:
            return

        input_dir = Path(self.input_dir_var.get().strip()).expanduser()
        output_dir = Path(self.output_dir_var.get().strip()).expanduser()
        mapping_value = self.mapping_file_var.get().strip()
        mapping_path = Path(mapping_value).expanduser() if mapping_value else DEFAULT_MAPPING_FILE

        if not input_dir.exists() or not input_dir.is_dir():
            self.messagebox.showerror("Errore", "La cartella input non esiste.")
            return

        if not mapping_path.exists() or not mapping_path.is_file():
            self.messagebox.showerror("Errore", "La banca dati di trascodifica non esiste.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self._save_state()
        self.clear_views()
        self._set_busy(True, "Esecuzione ETL in corso...")
        self.append_console(f"Input: {input_dir}", "header")
        self.append_console(f"Output: {output_dir}", "header")
        self.append_console(f"Trascodifica: {mapping_path}", "header")

        threading.Thread(target=self._run_pipeline_worker, args=(input_dir, output_dir, mapping_path), daemon=True).start()

    def _run_pipeline_worker(self, input_dir: Path, output_dir: Path, mapping_path: Path) -> None:
        try:
            try:
                from . import etl_bancadati as etl
                from . import validate_output as validate
            except ImportError:
                import etl_bancadati as etl
                import validate_output as validate

            etl.BANCADATI_DIR = input_dir
            etl.OUTPUT_DIR = output_dir
            etl.MAPPING_FILE = mapping_path
            validate.BANCADATI_DIR = input_dir
            validate.OUTPUT_DIR = output_dir

            self.append_console("Avvio preparazione file...", "info")
            with redirect_stdout(TkLogStream(self, "info")), redirect_stderr(TkLogStream(self, "error")):
                etl.main()

            self.root.after(0, lambda: self._load_anomalies_view(output_dir))

            self.append_console("Avvio validazione output...", "info")
            with redirect_stdout(TkLogStream(self, "info")), redirect_stderr(TkLogStream(self, "error")):
                validate.main()

            self.root.after(0, lambda: self._load_validation_view(output_dir))
            self.root.after(0, lambda: self.status_var.set("Elaborazione completata."))
            self.append_console("Pipeline ETL + validazione completata.", "success")
        except Exception:
            self.append_console(traceback.format_exc(), "error")
            self.root.after(0, lambda: self.status_var.set("Errore durante l'elaborazione."))
        finally:
            self.root.after(0, lambda: self._set_busy(False))

    def _load_anomalies_view(self, output_dir: Path) -> None:
        anomaly_log_path = output_dir / "_anomalie.tsv"
        unresolved_path = output_dir / "_prestazioni_non_ricavabili.tsv"
        excel_report_path = output_dir / "_anomalie_non_riconducibili_ssn.xlsx"

        issue_counts: Counter[str] = Counter()
        unresolved_rows: list[dict[str, str]] = []

        if anomaly_log_path.exists():
            with anomaly_log_path.open(encoding="utf-8", newline="") as handle:
                reader = csv.DictReader(handle, delimiter="\t")
                for row in reader:
                    issue_counts[row["issue"]] += 1

        if unresolved_path.exists():
            with unresolved_path.open(encoding="utf-8", newline="") as handle:
                reader = csv.DictReader(handle, delimiter="\t")
                unresolved_rows = list(reader)

        lines: list[str] = []
        lines.append(f"Report Excel unico: {excel_report_path}")
        lines.append(f"File input con prestazioni non riconducibili a SSN: {len(unresolved_rows)}")
        lines.append("")
        lines.append("Conteggio anomalie per tipo:")
        for issue, count in issue_counts.most_common():
            lines.append(f"- {issue}: {count}")

        if unresolved_rows:
            lines.append("")
            lines.append("Dettaglio anomalie residue:")
            for row in unresolved_rows:
                lines.append(
                    f"- {row['folder_name']}/{row['source_file']}: {row['prestazioni_non_ricavabili']}"
                )

        self._set_text(self.anomalies_text, "\n".join(lines).strip() + "\n")
        self.append_console(f"Report anomalie caricato: {excel_report_path}", "success")

    def _load_validation_view(self, output_dir: Path) -> None:
        summary_path = output_dir / "_validation_summary.txt"
        report_path = output_dir / "_validation_report.tsv"

        lines: list[str] = []
        if summary_path.exists():
            lines.append(summary_path.read_text(encoding="utf-8").strip())

        warnings_count = 0
        error_count = 0
        report_rows: list[dict[str, str]] = []
        if report_path.exists():
            with report_path.open(encoding="utf-8", newline="") as handle:
                reader = csv.DictReader(handle, delimiter="\t")
                report_rows = list(reader)
            warning_rows = [row for row in report_rows if row["severity"] == "warning"]
            error_rows = [row for row in report_rows if row["severity"] == "error"]
            warnings_count = len(warning_rows)
            error_count = len(error_rows)

            lines.append("")
            lines.append(f"Errori strutturali: {error_count}")
            lines.append(f"Warning semantici: {warnings_count}")

            if warning_rows:
                lines.append("")
                lines.append("Warning:")
                for row in warning_rows:
                    lines.append(f"- {row['output_file']}: {row['details']}")

        self._set_text(self.validation_text, "\n".join(lines).strip() + "\n")
        self.append_console(f"Validazione completata: errori={error_count}, warning={warnings_count}", "success" if error_count == 0 else "warn")

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

    def open_output_dir(self) -> None:
        self._open_path(Path(self.output_dir_var.get().strip()).expanduser())

    def open_anomaly_excel(self) -> None:
        self._open_path(Path(self.output_dir_var.get().strip()).expanduser() / "_anomalie_non_riconducibili_ssn.xlsx")

    def open_validation_summary(self) -> None:
        self._open_path(Path(self.output_dir_var.get().strip()).expanduser() / "_validation_summary.txt")

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

    def _save_state(self) -> None:
        payload = {
            "version": self.STATE_VERSION,
            "input_dir": self.input_dir_var.get().strip(),
            "output_dir": self.output_dir_var.get().strip(),
            "mapping_file": self.mapping_file_var.get().strip(),
        }
        STATE_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _load_state(self) -> None:
        if not STATE_FILE.exists():
            return
        try:
            payload = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        input_dir = payload.get("input_dir")
        output_dir = payload.get("output_dir")
        mapping_file = payload.get("mapping_file")
        if input_dir:
            self.input_dir_var.set(input_dir)
        if output_dir:
            self.output_dir_var.set(output_dir)
        if mapping_file:
            self.mapping_file_var.set(mapping_file)


def main() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    except ModuleNotFoundError as exc:
        raise SystemExit(f"Tkinter non disponibile: {exc}")

    root = tk.Tk()
    app = SpecialisticaGuiApp(root, tk, ttk, filedialog, messagebox, scrolledtext)
    root.mainloop()


if __name__ == "__main__":
    main()
