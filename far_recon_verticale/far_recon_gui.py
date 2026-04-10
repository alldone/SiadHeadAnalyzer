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


PROJECT_ROOT = Path(__file__).resolve().parent
STATE_FILE = PROJECT_ROOT / ".far_recon_gui_state.json"


class FarReconApp:
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
            self.root.title("FAR Reconciliation NSIS / SISR")
            self.root.geometry("1320x860")
            self.root.minsize(1100, 740)

        self.nsis_dir_var = self.tk.StringVar(value="")
        self.nuovi_dir_var = self.tk.StringVar(value="")
        self.scarti_dir_var = self.tk.StringVar(value="")
        self.output_file_var = self.tk.StringVar(value="")
        self.status_var = self.tk.StringVar(
            value="Seleziona le cartelle acquisiti_NSIS e/o nuovi_FLUSSI, poi avvia il report."
        )
        self.spinner_var = self.tk.StringVar(value="")

        self._busy = False
        self._spinner_index = 0
        self._spinner_job: str | None = None
        self._last_output: str | None = None

        self._build_ui()
        self._load_state()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

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

        # Ribbon
        ribbon = self.ttk.Frame(frame)
        ribbon.grid(row=0, column=0, sticky="ew")
        ribbon.columnconfigure(0, weight=1)
        ribbon.columnconfigure(1, weight=1)
        ribbon.columnconfigure(2, weight=1)

        self._build_ribbon_group(
            ribbon, 0, "Percorsi",
            [
                ("Acquisiti NSIS", self.choose_nsis_dir),
                ("Nuovi FLUSSI", self.choose_nuovi_dir),
                ("Scarti SISR", self.choose_scarti_dir),
                ("File output", self.choose_output_file),
            ],
        )
        self._build_ribbon_group(
            ribbon, 1, "Esecuzione",
            [("Genera report", self.run_report), ("Pulisci log", self.clear_views)],
        )
        self._build_ribbon_group(
            ribbon, 2, "Report",
            [("Apri Excel", self.open_output_file), ("Apri cartella", self.open_output_dir)],
        )

        # Percorsi
        path_frame = self.ttk.LabelFrame(frame, text="Configurazione percorsi", padding=12)
        path_frame.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        path_frame.columnconfigure(1, weight=1)

        self._add_path_row(path_frame, 0, "Acquisiti NSIS", self.nsis_dir_var, self.choose_nsis_dir)
        self._add_path_row(path_frame, 1, "Nuovi FLUSSI", self.nuovi_dir_var, self.choose_nuovi_dir)
        self._add_path_row(path_frame, 2, "Scarti SISR (opz.)", self.scarti_dir_var, self.choose_scarti_dir)
        self._add_path_row(path_frame, 3, "File Excel output", self.output_file_var, self.choose_output_file, file_mode=True)

        hint = (
            "Acquisiti NSIS: cartella con XML/zip gia' acquisiti. "
            "Nuovi FLUSSI: cartella con XML del nuovo invio (scarti in sottocartella 'scarti/'). "
            "Se entrambe presenti viene calcolata la proiezione post-upload con chiavi NSIS § 4.3. "
            "La cartella scarti SISR viene dedotta automaticamente da nuovi_FLUSSI/scarti/ se non specificata."
        )
        self.ttk.Label(frame, text=hint, wraplength=1260, justify="left").grid(row=2, column=0, sticky="ew", pady=(0, 8))

        # Notebook
        self.notebook = self.ttk.Notebook(frame)
        self.notebook.grid(row=3, column=0, sticky="nsew")

        self.console_text = self._add_text_tab("Output")
        self.validation_text = self._add_text_tab("Riepilogo")

        # Status bar
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
        file_menu.add_command(label="Acquisiti NSIS...", command=self.choose_nsis_dir)
        file_menu.add_command(label="Nuovi FLUSSI...", command=self.choose_nuovi_dir)
        file_menu.add_command(label="Scarti SISR...", command=self.choose_scarti_dir)
        file_menu.add_command(label="File output...", command=self.choose_output_file)
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

    def _add_path_row(self, parent: Any, row: int, label: str, variable: Any, command: Any, file_mode: bool = False) -> None:
        self.ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        self.ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        self.ttk.Button(parent, text="Sfoglia...", command=command).grid(row=row, column=2, sticky="e", pady=4)

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

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    @staticmethod
    def _next_output_path(base_dir: Path, stem: str = "conteggio_FAR_recon", ext: str = ".xlsx") -> Path:
        candidate = base_dir / f"{stem}{ext}"
        if not candidate.exists():
            return candidate
        n = 1
        while True:
            candidate = base_dir / f"{stem}_{n}{ext}"
            if not candidate.exists():
                return candidate
            n += 1

    def choose_nsis_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.nsis_dir_var.get() or str(Path.home()))
        if path:
            self.nsis_dir_var.set(path)
            if not self.output_file_var.get().strip():
                self.output_file_var.set(str(self._next_output_path(Path(path).parent)))
            self._save_state()

    def choose_nuovi_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.nuovi_dir_var.get() or str(Path.home()))
        if path:
            self.nuovi_dir_var.set(path)
            # auto-detect scarti subfolder
            scarti_candidate = os.path.join(path, "scarti")
            if os.path.isdir(scarti_candidate) and not self.scarti_dir_var.get().strip():
                self.scarti_dir_var.set(scarti_candidate)
            if not self.output_file_var.get().strip():
                self.output_file_var.set(str(self._next_output_path(Path(path).parent)))
            self._save_state()

    def choose_scarti_dir(self) -> None:
        path = self.filedialog.askdirectory(initialdir=self.scarti_dir_var.get() or str(Path.home()))
        if path:
            self.scarti_dir_var.set(path)
            self._save_state()

    def choose_output_file(self) -> None:
        current = self.output_file_var.get().strip()
        init_dir = str(Path(current).parent) if current else str(Path.home())
        path = self.filedialog.asksaveasfilename(
            initialdir=init_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.output_file_var.set(path)
            self._save_state()

    def clear_views(self) -> None:
        for widget in (self.console_text, self.validation_text):
            self._set_text(widget, "")
        self.status_var.set("Log ripuliti.")

    def run_report(self) -> None:
        if self._busy:
            return

        nsis = self.nsis_dir_var.get().strip()
        nuovi = self.nuovi_dir_var.get().strip()
        output = self.output_file_var.get().strip()

        if not nsis and not nuovi:
            self.messagebox.showerror("Errore", "Seleziona almeno una cartella (NSIS o Nuovi FLUSSI).")
            return
        if not output:
            self.messagebox.showerror("Errore", "Indica il percorso del file Excel di output.")
            return

        # Incrementale se esiste
        out_path = Path(output)
        if out_path.exists():
            out_path = self._next_output_path(out_path.parent)
            output = str(out_path)
            self.output_file_var.set(output)

        scarti = self.scarti_dir_var.get().strip() or None

        self._save_state()
        self.clear_views()
        self._set_busy(True, "Elaborazione in corso...")

        self.append_console("=== FAR Reconciliation NSIS / SISR ===", "header")
        if nsis:
            self.append_console(f"Acquisiti NSIS: {nsis}", "header")
        if nuovi:
            self.append_console(f"Nuovi FLUSSI:   {nuovi}", "header")
        if scarti:
            self.append_console(f"Scarti SISR:    {scarti}", "header")
        self.append_console(f"Output:         {output}", "header")
        self.append_console("", "info")

        threading.Thread(
            target=self._run_worker,
            args=(nsis or None, nuovi or None, scarti, output),
            daemon=True,
        ).start()

    def _run_worker(self, nsis: str | None, nuovi: str | None, scarti: str | None, output: str) -> None:
        try:
            try:
                from .far_recon_core import generate_report
            except ImportError:
                from far_recon_core import generate_report

            result = generate_report(
                nsis_dir=nsis,
                nuovi_dir=nuovi,
                scarti_dir=scarti,
                output_path=output,
                log=lambda msg: self.append_console(msg, "info"),
            )
            self._last_output = result
            self.root.after(0, lambda: self.status_var.set(f"Report generato: {result}"))
            self.append_console("", "info")
            self.append_console("Elaborazione completata con successo.", "success")

            # Riepilogo
            summary = self._build_summary(nsis, nuovi, scarti)
            self.root.after(0, lambda: self._set_text(self.validation_text, summary))
        except Exception:
            self.append_console(traceback.format_exc(), "error")
            self.root.after(0, lambda: self.status_var.set("Errore durante l'elaborazione."))
        finally:
            self.root.after(0, lambda: self._set_busy(False))

    def _build_summary(self, nsis: str | None, nuovi: str | None, scarti: str | None) -> str:
        lines = ["=== Riepilogo elaborazione FAR ===\n"]
        if nsis:
            lines.append(f"Acquisiti NSIS: {nsis}")
        if nuovi:
            lines.append(f"Nuovi FLUSSI: {nuovi}")
        if scarti:
            lines.append(f"Scarti SISR: {scarti}")
        lines.append(f"\nFile generato: {self._last_output}")
        lines.append("\nSheet presenti nel workbook:")
        lines.append("  - acquisiti_NSIS (se cartella NSIS fornita)")
        lines.append("  - nuovi_FLUSSI (con colonna netto scarti)")
        lines.append("  - scarti_nuovi_FLUSSI")
        lines.append("  - delta_validi")
        lines.append("  - scarti_per_trimestre (logica Nicoletta: Azienda x Trimestre x Tracciato)")
        if nsis and nuovi:
            lines.append("  - proiezione_post_upload (chiavi NSIS § 4.3)")
        return "\n".join(lines)

    # ------------------------------------------------------------------
    # Report opening
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Console helpers
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Busy / spinner
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # State persistence
    # ------------------------------------------------------------------

    def _save_state(self) -> None:
        payload = {
            "version": self.STATE_VERSION,
            "nsis_dir": self.nsis_dir_var.get().strip(),
            "nuovi_dir": self.nuovi_dir_var.get().strip(),
            "scarti_dir": self.scarti_dir_var.get().strip(),
            "output_file": self.output_file_var.get().strip(),
        }
        STATE_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _load_state(self) -> None:
        if not STATE_FILE.exists():
            return
        try:
            payload = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        for key, var in [
            ("nsis_dir", self.nsis_dir_var),
            ("nuovi_dir", self.nuovi_dir_var),
            ("scarti_dir", self.scarti_dir_var),
            ("output_file", self.output_file_var),
        ]:
            val = payload.get(key)
            if val:
                var.set(val)


def main() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    except ModuleNotFoundError as exc:
        raise SystemExit(f"Tkinter non disponibile: {exc}")

    root = tk.Tk()
    FarReconApp(root, tk, ttk, filedialog, messagebox, scrolledtext)
    root.mainloop()


if __name__ == "__main__":
    main()
