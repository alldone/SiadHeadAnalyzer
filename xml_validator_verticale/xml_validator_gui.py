#!/usr/bin/env python3

from __future__ import annotations

from pathlib import Path
from queue import Empty, Queue
import threading
from typing import Any

try:
    from .xml_validator_core import ValidationIssue, ValidationResult, validate_xml_without_namespaces
except ImportError:
    from xml_validator_core import ValidationIssue, ValidationResult, validate_xml_without_namespaces


WorkerMessage = tuple[str, object]


class XmlValidatorApp:
    """GUI generica e non bloccante per la validazione XML/XSD."""

    POLL_INTERVAL_MS = 50
    MAX_MESSAGES_PER_POLL = 250

    def __init__(
        self,
        root: Any,
        tk_module: Any,
        ttk_module: Any,
        filedialog_module: Any,
        messagebox_module: Any,
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

        if not self.embed_mode:
            self.root.title("Validatore XML/XSD")
            self.root.geometry("1100x720")
            self.root.minsize(820, 540)

        self.xsd_path_var = self.tk.StringVar(value="")
        self.xml_path_var = self.tk.StringVar(value="")
        self.status_var = self.tk.StringVar(value="Seleziona un file XSD e un file XML da validare.")
        self.result_title_var = self.tk.StringVar(value="Errori di validazione")

        self._busy = False
        self._issue_count = 0
        self._messages: Queue[WorkerMessage] = Queue()
        self._poll_job: str | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        if not self.embed_mode:
            self._build_menu()

        frame = self.ttk.Frame(self.parent, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        path_frame = self.ttk.LabelFrame(frame, text="File da validare", padding=12)
        path_frame.grid(row=0, column=0, sticky="ew")
        path_frame.columnconfigure(1, weight=1)

        self.xsd_browse_button = self._add_path_row(
            path_frame,
            0,
            "Schema XSD",
            self.xsd_path_var,
            self.choose_xsd,
        )
        self.xml_browse_button = self._add_path_row(
            path_frame,
            1,
            "Documento XML",
            self.xml_path_var,
            self.choose_xml,
        )

        action_frame = self.ttk.Frame(frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        action_frame.columnconfigure(2, weight=1)

        self.validate_button = self.ttk.Button(action_frame, text="Valida XML", command=self.start_validation)
        self.validate_button.grid(row=0, column=0, sticky="w")
        self.clear_button = self.ttk.Button(action_frame, text="Pulisci risultati", command=self.clear_results)
        self.clear_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.ttk.Label(
            action_frame,
            text=(
                "I namespace dell'XML vengono ignorati su una copia temporanea; "
                "i file originali non vengono modificati."
            ),
            justify="right",
        ).grid(row=0, column=2, sticky="e")

        results_frame = self.ttk.LabelFrame(frame, textvariable=self.result_title_var, padding=8)
        results_frame.grid(row=2, column=0, sticky="nsew")
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

        self.error_listbox = self.tk.Listbox(
            results_frame,
            activestyle="dotbox",
            exportselection=False,
            selectmode="browse",
        )
        self.error_listbox.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar = self.ttk.Scrollbar(
            results_frame,
            orient="vertical",
            command=self.error_listbox.yview,
        )
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar = self.ttk.Scrollbar(
            results_frame,
            orient="horizontal",
            command=self.error_listbox.xview,
        )
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.error_listbox.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )

        status_frame = self.ttk.Frame(frame)
        status_frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        status_frame.columnconfigure(1, weight=1)
        self.progress = self.ttk.Progressbar(status_frame, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.ttk.Label(status_frame, textvariable=self.status_var, anchor="w").grid(
            row=0,
            column=1,
            sticky="ew",
        )

    def _build_menu(self) -> None:
        menu = self.tk.Menu(self.root)
        file_menu = self.tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Seleziona XSD...", command=self.choose_xsd)
        file_menu.add_command(label="Seleziona XML...", command=self.choose_xml)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        validation_menu = self.tk.Menu(menu, tearoff=False)
        validation_menu.add_command(label="Valida XML", command=self.start_validation)
        validation_menu.add_command(label="Pulisci risultati", command=self.clear_results)
        menu.add_cascade(label="Validazione", menu=validation_menu)
        self.root.config(menu=menu)

    def _add_path_row(self, parent: Any, row: int, label: str, variable: Any, command: Any) -> Any:
        self.ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        self.ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        button = self.ttk.Button(parent, text="Sfoglia...", command=command)
        button.grid(row=row, column=2, sticky="e", pady=4)
        return button

    def _initial_directory(self, current_path: str) -> str:
        candidate = Path(current_path).expanduser()
        if candidate.is_file():
            return str(candidate.parent)
        return str(Path.home())

    def choose_xsd(self) -> None:
        path = self.filedialog.askopenfilename(
            title="Seleziona lo schema XSD",
            initialdir=self._initial_directory(self.xsd_path_var.get()),
            filetypes=[("XML Schema", "*.xsd"), ("Tutti i file", "*.*")],
        )
        if path:
            self.xsd_path_var.set(path)

    def choose_xml(self) -> None:
        path = self.filedialog.askopenfilename(
            title="Seleziona il documento XML",
            initialdir=self._initial_directory(self.xml_path_var.get()),
            filetypes=[("File XML", "*.xml"), ("Tutti i file", "*.*")],
        )
        if path:
            self.xml_path_var.set(path)

    def clear_results(self) -> None:
        if self._busy:
            return
        self.error_listbox.delete(0, "end")
        self._issue_count = 0
        self.result_title_var.set("Errori di validazione")
        self.status_var.set("Risultati cancellati. Seleziona i file e avvia una nuova validazione.")

    def start_validation(self) -> None:
        if self._busy:
            return

        xsd_path = Path(self.xsd_path_var.get().strip()).expanduser()
        xml_path = Path(self.xml_path_var.get().strip()).expanduser()
        if not xsd_path.is_file():
            self.messagebox.showerror("Errore", "Seleziona un file XSD esistente.")
            return
        if not xml_path.is_file():
            self.messagebox.showerror("Errore", "Seleziona un file XML esistente.")
            return

        self.error_listbox.delete(0, "end")
        self._issue_count = 0
        self.result_title_var.set("Errori di validazione (0)")
        self._set_busy(True, f"Validazione in corso: {xml_path.name}")

        # La coda separa completamente il worker dalla GUI Tk, che può essere
        # aggiornata solo dal thread principale.
        self._messages = Queue()
        threading.Thread(
            target=self._validation_worker,
            args=(xml_path, xsd_path),
            name="xml-xsd-validator",
            daemon=True,
        ).start()
        self._schedule_poll()

    def _validation_worker(self, xml_path: Path, xsd_path: Path) -> None:
        try:
            result = validate_xml_without_namespaces(
                xml_path,
                xsd_path,
                on_issue=lambda issue: self._messages.put(("issue", issue)),
            )
            self._messages.put(("done", result))
        except Exception as exc:
            self._messages.put(("error", str(exc)))

    def _schedule_poll(self) -> None:
        if self._poll_job is None:
            self._poll_job = self.root.after(self.POLL_INTERVAL_MS, self._poll_messages)

    def _poll_messages(self) -> None:
        self._poll_job = None
        listbox_rows: list[str] = []
        terminal_message: WorkerMessage | None = None

        for _ in range(self.MAX_MESSAGES_PER_POLL):
            try:
                kind, payload = self._messages.get_nowait()
            except Empty:
                break

            if kind == "issue" and isinstance(payload, ValidationIssue):
                self._issue_count += 1
                listbox_rows.append(payload.display_text(self._issue_count))
            elif kind in {"done", "error"}:
                terminal_message = (kind, payload)
                break

        if listbox_rows:
            self.error_listbox.insert("end", *listbox_rows)
            self.error_listbox.see("end")
            self.result_title_var.set(f"Errori di validazione ({self._issue_count})")
            self.status_var.set(f"Validazione in corso: {self._issue_count} errore/i rilevati...")

        if terminal_message is None:
            self._schedule_poll()
            return

        kind, payload = terminal_message
        if kind == "done" and isinstance(payload, ValidationResult):
            self._finish_success(payload)
        else:
            self._finish_error(str(payload))

    def _finish_success(self, result: ValidationResult) -> None:
        self._set_busy(False)
        if result.issues:
            self.status_var.set(f"Validazione completata: {len(result.issues)} errore/i rilevati.")
        else:
            self.error_listbox.insert("end", "Nessun errore: il documento XML è conforme allo schema XSD.")
            self.status_var.set("Validazione completata: XML conforme allo XSD.")

        if result.namespace_was_present:
            self.status_var.set(
                f"{self.status_var.get()} Namespace XML ignorati: {', '.join(result.source_namespaces)}"
            )

    def _finish_error(self, message: str) -> None:
        self._set_busy(False)
        self.error_listbox.insert("end", f"Errore durante la validazione: {message}")
        self.status_var.set("Validazione non completata.")
        self.messagebox.showerror("Errore di validazione", message)

    def _set_busy(self, busy: bool, status: str | None = None) -> None:
        self._busy = busy
        if status is not None:
            self.status_var.set(status)
        state = "disabled" if busy else "normal"
        self.validate_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.xsd_browse_button.configure(state=state)
        self.xml_browse_button.configure(state=state)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()


def main() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ModuleNotFoundError as exc:
        raise SystemExit(f"Tkinter non disponibile: {exc}")

    root = tk.Tk()
    XmlValidatorApp(root, tk, ttk, filedialog, messagebox)
    root.mainloop()


if __name__ == "__main__":
    main()
