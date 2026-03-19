#!/usr/bin/env python3

from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys
from typing import Any

from siad_report_gui import SiadReportApp
from specialistica_verticale.specialistica_gui import SpecialisticaGuiApp


PROJECT_ROOT = Path(__file__).resolve().parent
README_PATH = PROJECT_ROOT / "README.md"
SPECIALISTICA_NOTE_PATH = PROJECT_ROOT / "specialistica_verticale" / "NOTE_ETL_SPECIALISTICA.md"


class ToolSuiteApp:
    def __init__(
        self,
        root: Any,
        tk_module: Any,
        ttk_module: Any,
        filedialog_module: Any,
        messagebox_module: Any,
        scrolledtext_module: Any,
    ) -> None:
        self.root = root
        self.tk = tk_module
        self.ttk = ttk_module
        self.filedialog = filedialog_module
        self.messagebox = messagebox_module
        self.scrolledtext = scrolledtext_module

        self.root.title("SiadHeadAnalyzer")
        self.root.geometry("1460x920")
        self.root.minsize(1240, 820)

        self.status_var = self.tk.StringVar(value="Seleziona il verticale da usare dal ribbon o dalle tab.")
        self.siad_app: SiadReportApp | None = None
        self.specialistica_app: SpecialisticaGuiApp | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        self._build_menu()

        style = self.ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        frame = self.ttk.Frame(self.root, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ribbon = self.ttk.Frame(frame)
        ribbon.grid(row=0, column=0, sticky="ew")
        ribbon.columnconfigure(0, weight=1)
        ribbon.columnconfigure(1, weight=1)
        ribbon.columnconfigure(2, weight=1)

        self._build_ribbon_group(
            ribbon,
            0,
            "Verticali",
            [
                ("Home", self.show_home),
                ("SIAD", self.show_siad),
                ("Specialistica", self.show_specialistica),
            ],
        )
        self._build_ribbon_group(
            ribbon,
            1,
            "Avvio Rapido",
            [
                ("Apri SIAD", self.show_siad),
                ("Apri Specialistica", self.show_specialistica),
            ],
        )
        self._build_ribbon_group(
            ribbon,
            2,
            "Documentazione",
            [
                ("README", lambda: self._open_path(README_PATH)),
                ("Note ETL", lambda: self._open_path(SPECIALISTICA_NOTE_PATH)),
            ],
        )

        self.notebook = self.ttk.Notebook(frame)
        self.notebook.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self.home_tab = self.ttk.Frame(self.notebook, padding=20)
        self.siad_tab = self.ttk.Frame(self.notebook)
        self.specialistica_tab = self.ttk.Frame(self.notebook)
        self.notebook.add(self.home_tab, text="Home")
        self.notebook.add(self.siad_tab, text="SIAD")
        self.notebook.add(self.specialistica_tab, text="Specialistica")

        self._build_home_tab()

        status_bar = self.ttk.Frame(frame)
        status_bar.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        status_bar.columnconfigure(0, weight=1)
        self.ttk.Label(status_bar, textvariable=self.status_var, anchor="w").grid(row=0, column=0, sticky="ew")

    def _build_menu(self) -> None:
        menu = self.tk.Menu(self.root)

        file_menu = self.tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Home", command=self.show_home)
        file_menu.add_command(label="SIAD", command=self.show_siad)
        file_menu.add_command(label="Specialistica", command=self.show_specialistica)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        docs_menu = self.tk.Menu(menu, tearoff=False)
        docs_menu.add_command(label="README", command=lambda: self._open_path(README_PATH))
        docs_menu.add_command(label="Note ETL Specialistica", command=lambda: self._open_path(SPECIALISTICA_NOTE_PATH))
        menu.add_cascade(label="Documentazione", menu=docs_menu)

        self.root.config(menu=menu)

    def _build_ribbon_group(self, parent: Any, column: int, title: str, buttons: list[tuple[str, Any]]) -> None:
        group = self.ttk.LabelFrame(parent, text=title, padding=10)
        group.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column < 2 else 0)
        for index, (label, command) in enumerate(buttons):
            button = self.ttk.Button(group, text=label, command=command)
            button.grid(row=0, column=index, padx=(0, 8) if index < len(buttons) - 1 else 0, pady=2, ipadx=8, ipady=10)

    def _build_home_tab(self) -> None:
        self.home_tab.columnconfigure(0, weight=1)
        self.home_tab.columnconfigure(1, weight=1)

        intro = (
            "Seleziona il verticale operativo da usare. "
            "La codebase resta separata: SIAD e Specialistica condividono solo questo launcher."
        )
        self.ttk.Label(self.home_tab, text=intro, wraplength=1200, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 18)
        )

        self._build_tool_card(
            self.home_tab,
            1,
            0,
            "SIAD",
            "Analisi XML, generazione report Excel e validazione XSD dei tracciati SIAD.",
            self.show_siad,
        )
        self._build_tool_card(
            self.home_tab,
            1,
            1,
            "Specialistica",
            "Preparazione ETL, trascodifica prestazioni, report anomalie e validazione output.",
            self.show_specialistica,
        )

    def _build_tool_card(
        self,
        parent: Any,
        row: int,
        column: int,
        title: str,
        description: str,
        command: Any,
    ) -> None:
        card = self.ttk.LabelFrame(parent, text=title, padding=16)
        card.grid(row=row, column=column, sticky="nsew", padx=(0, 12) if column == 0 else 0)
        card.columnconfigure(0, weight=1)
        self.ttk.Label(card, text=description, wraplength=520, justify="left").grid(row=0, column=0, sticky="w")
        self.ttk.Button(card, text=f"Apri {title}", command=command).grid(row=1, column=0, sticky="w", pady=(18, 0))

    def show_home(self) -> None:
        self.notebook.select(self.home_tab)
        self.status_var.set("Home attiva.")

    def show_siad(self) -> None:
        self._ensure_app_loaded("siad")
        self.notebook.select(self.siad_tab)
        self.status_var.set("Verticale SIAD attivo.")

    def show_specialistica(self) -> None:
        self._ensure_app_loaded("specialistica")
        self.notebook.select(self.specialistica_tab)
        self.status_var.set("Verticale Specialistica attivo.")

    def _on_tab_changed(self, _event: Any) -> None:
        selected = self.notebook.select()
        if selected == str(self.siad_tab):
            self._ensure_app_loaded("siad")
            self.status_var.set("Verticale SIAD attivo.")
        elif selected == str(self.specialistica_tab):
            self._ensure_app_loaded("specialistica")
            self.status_var.set("Verticale Specialistica attivo.")
        else:
            self.status_var.set("Home attiva.")

    def _ensure_app_loaded(self, tool_name: str) -> None:
        if tool_name == "siad" and self.siad_app is None:
            self.siad_app = SiadReportApp(
                self.root,
                self.tk,
                self.ttk,
                self.filedialog,
                self.messagebox,
                parent=self.siad_tab,
                embed_mode=True,
            )
            return

        if tool_name == "specialistica" and self.specialistica_app is None:
            self.specialistica_app = SpecialisticaGuiApp(
                self.root,
                self.tk,
                self.ttk,
                self.filedialog,
                self.messagebox,
                self.scrolledtext,
                parent=self.specialistica_tab,
                embed_mode=True,
            )

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


def main() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "Tkinter non disponibile in questa installazione Python. "
            "Per usare la GUI installa una build Python con supporto Tk."
        ) from exc

    root = tk.Tk()
    ToolSuiteApp(root, tk, ttk, filedialog, messagebox, scrolledtext)
    root.mainloop()


if __name__ == "__main__":
    main()
