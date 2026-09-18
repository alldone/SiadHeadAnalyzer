#!/usr/bin/env python3
"""Motore del verticale Telemedicina/PNT.

La sorgente viene letta per nome delle sole colonne effettivamente utilizzate;
le colonne aggiuntive vengono ignorate. La quantità viene calcolata come
conteggio degli ID paziente, replicando il criterio della pivot.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import warnings
import zipfile
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

try:
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
except ImportError as exc:  # pragma: no cover - messaggio per ambienti nuovi
    raise SystemExit(
        "Dipendenza mancante: installare openpyxl (python -m pip install openpyxl)."
    ) from exc


IGNORED_DIRS = {".git", ".codex", "outputs", "__pycache__"}
WORKSHEET_EXTENSION_RE = re.compile(
    rb"<extLst(?:\s[^>]*)?>.*?</extLst>", re.DOTALL
)
XR_NAMESPACE = b"http://schemas.microsoft.com/office/spreadsheetml/2014/revision"

# openpyxl avvisa che le estensioni x14 saranno rimosse. Più avanti vengono
# reinserite byte-per-byte dal modello, così restano disponibili i menu a tendina.
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", text).strip()


def text_key(value: Any) -> str:
    return clean_text(value).casefold()


def code_value(value: Any, width: int) -> str:
    """Normalizza codici numerici conservando/ricreando gli zeri iniziali."""
    if value is None or clean_text(value) == "":
        raise ValueError("codice vuoto")
    if isinstance(value, bool):
        raise ValueError(f"codice booleano non valido: {value!r}")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not value.is_integer():
            raise ValueError(f"codice non intero: {value!r}")
        digits = str(int(value))
    else:
        text = clean_text(value)
        if re.fullmatch(r"\d+(?:\.0+)?", text):
            digits = text.split(".", 1)[0]
        else:
            digits = re.sub(r"\D", "", text)
            if not digits:
                raise ValueError(f"codice non numerico: {value!r}")
    if len(digits) > width:
        raise ValueError(f"codice {digits!r} più lungo di {width} cifre")
    return digits.zfill(width)


def read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def candidate_xlsx_files(input_dir: Path) -> Iterable[Path]:
    for path in sorted(input_dir.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        relative_parts = path.relative_to(input_dir).parts[:-1]
        if any(part in IGNORED_DIRS for part in relative_parts):
            continue
        yield path


def workbook_sheet_names(path: Path) -> set[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return set(workbook.sheetnames)
    finally:
        workbook.close()


def discover_workbook(
    input_dir: Path,
    required_sheets: set[str],
    explicit_path: Path | None,
    role: str,
) -> Path:
    if explicit_path is not None:
        path = explicit_path.resolve()
        missing = required_sheets - workbook_sheet_names(path)
        if missing:
            raise RuntimeError(
                f"Il file indicato come {role} non contiene i fogli: "
                + ", ".join(sorted(missing))
            )
        return path

    matches: list[Path] = []
    for path in candidate_xlsx_files(input_dir):
        try:
            if required_sheets.issubset(workbook_sheet_names(path)):
                matches.append(path.resolve())
        except Exception as exc:
            print(f"AVVISO: impossibile leggere {path}: {exc}", file=sys.stderr)

    if not matches:
        raise RuntimeError(
            f"Nessun file {role} trovato. Fogli richiesti: "
            + ", ".join(sorted(required_sheets))
        )
    if len(matches) > 1:
        paths = "\n  - ".join(str(path) for path in matches)
        raise RuntimeError(
            f"Trovati più file candidati per {role}; usare l'opzione esplicita:\n  - {paths}"
        )
    return matches[0]


def find_header_row(worksheet: Any, aliases: dict[str, list[str]]) -> tuple[int, dict[str, int]]:
    normalized_aliases = {
        field: [text_key(alias) for alias in field_aliases]
        for field, field_aliases in aliases.items()
    }
    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=min(25, worksheet.max_row), values_only=True),
        start=1,
    ):
        positions: dict[str, int] = {}
        by_name: dict[str, int] = {}
        for column_number, value in enumerate(row, start=1):
            key = text_key(value)
            if key and key not in by_name:
                by_name[key] = column_number
        for field, possible_names in normalized_aliases.items():
            for possible_name in possible_names:
                if possible_name in by_name:
                    positions[field] = by_name[possible_name]
                    break
        if len(positions) == len(aliases):
            return row_number, positions
    missing = ", ".join(aliases)
    raise RuntimeError(
        f"Intestazioni sorgente non trovate nei primi 25 record: {missing}"
    )


@dataclass(frozen=True)
class SourceValidation:
    valid: bool
    source_path: Path
    sheet_name: str | None
    header_row: int | None
    columns: dict[str, int]
    missing_fields: tuple[str, ...]
    error: str | None = None

    def message(self, aliases: dict[str, list[str]]) -> str:
        if self.valid:
            used = ", ".join(aliases[field][0] for field in aliases)
            return (
                "File sorgente conforme.\n"
                f"Foglio: {self.sheet_name}\n"
                f"Riga intestazioni: {self.header_row}\n"
                f"Colonne utilizzate: {used}\n"
                "Le eventuali colonne aggiuntive non vengono considerate."
            )
        if self.error:
            return self.error
        details = []
        for field in self.missing_fields:
            accepted = aliases[field]
            details.append("- " + " oppure ".join(repr(name) for name in accepted))
        where = ""
        if self.sheet_name and self.header_row:
            where = f"\nMiglior candidato: foglio {self.sheet_name!r}, riga {self.header_row}."
        return (
            "File sorgente non conforme: mancano le seguenti colonne utilizzate:\n"
            + "\n".join(details)
            + where
            + "\nLe colonne non utilizzate dal verticale sono ignorate."
        )


class SourceValidationError(RuntimeError):
    def __init__(self, validation: SourceValidation, aliases: dict[str, list[str]]) -> None:
        self.validation = validation
        super().__init__(validation.message(aliases))


class MappingError(RuntimeError):
    """Uno o più valori sorgente richiedono una traduzione in configurazione."""


def _best_header_candidate(
    worksheet: Any, aliases: dict[str, list[str]]
) -> tuple[int | None, dict[str, int]]:
    normalized_aliases = {
        field: [text_key(alias) for alias in field_aliases]
        for field, field_aliases in aliases.items()
    }
    best_row: int | None = None
    best_positions: dict[str, int] = {}
    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=min(25, worksheet.max_row), values_only=True),
        start=1,
    ):
        by_name: dict[str, int] = {}
        for column_number, value in enumerate(row, start=1):
            key = text_key(value)
            if key and key not in by_name:
                by_name[key] = column_number
        positions: dict[str, int] = {}
        for field, possible_names in normalized_aliases.items():
            for possible_name in possible_names:
                if possible_name in by_name:
                    positions[field] = by_name[possible_name]
                    break
        if len(positions) > len(best_positions):
            best_row = row_number
            best_positions = positions
        if len(positions) == len(aliases):
            return row_number, positions
    return best_row, best_positions


def validate_source_workbook(
    source_path: str | Path, cfg: dict[str, Any] | None = None
) -> SourceValidation:
    """Valida esclusivamente le colonne realmente usate dal verticale.

    Il foglio configurato viene provato per primo. Se il suo nome cambia, il
    motore cerca automaticamente un altro foglio con tutte le intestazioni.
    """
    path = Path(source_path).expanduser().resolve()
    if not path.is_file():
        return SourceValidation(False, path, None, None, {}, (), f"File non trovato: {path}")
    if path.suffix.casefold() not in {".xlsx", ".xlsm"}:
        return SourceValidation(
            False,
            path,
            None,
            None,
            {},
            (),
            "Formato sorgente non supportato: selezionare un file .xlsx o .xlsm.",
        )

    active_cfg = cfg or load_default_config()
    aliases = active_cfg["source_columns"]
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        return SourceValidation(
            False, path, None, None, {}, (), f"Impossibile leggere il file Excel: {exc}"
        )

    try:
        preferred = active_cfg.get("source_sheet")
        ordered_names = list(workbook.sheetnames)
        if preferred in ordered_names:
            ordered_names.remove(preferred)
            ordered_names.insert(0, preferred)

        best_sheet: str | None = None
        best_row: int | None = None
        best_positions: dict[str, int] = {}
        for sheet_name in ordered_names:
            row_number, positions = _best_header_candidate(workbook[sheet_name], aliases)
            if len(positions) > len(best_positions):
                best_sheet = sheet_name
                best_row = row_number
                best_positions = positions
            if len(positions) == len(aliases):
                return SourceValidation(
                    True, path, sheet_name, row_number, positions, ()
                )

        missing = tuple(field for field in aliases if field not in best_positions)
        return SourceValidation(
            False, path, best_sheet, best_row, best_positions, missing
        )
    finally:
        workbook.close()


@dataclass(frozen=True)
class LookupValue:
    code: str
    display: str
    concatenated: str
    order: int


@dataclass
class LookupTable:
    by_code: dict[str, LookupValue]
    by_display: dict[str, LookupValue]


def load_lookup(worksheet: Any, code_width: int | None = None) -> LookupTable:
    header_row = None
    code_column = display_column = concatenated_column = None
    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=min(20, worksheet.max_row), values_only=True),
        start=1,
    ):
        headers = {text_key(value): index for index, value in enumerate(row, start=1)}
        if "code" in headers and "display" in headers:
            header_row = row_number
            code_column = headers["code"]
            display_column = headers["display"]
            concatenated_column = headers.get("concatenazione")
            break
    if header_row is None or code_column is None or display_column is None:
        raise RuntimeError(
            f"Il foglio {worksheet.title!r} non contiene le colonne Code e Display."
        )

    by_code: dict[str, LookupValue] = {}
    by_display: dict[str, LookupValue] = {}
    order = 0
    for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
        raw_code = row[code_column - 1]
        raw_display = row[display_column - 1]
        if raw_code is None or raw_display is None:
            continue
        code = code_value(raw_code, code_width) if code_width else clean_text(raw_code)
        display = clean_text(raw_display)
        raw_concatenated = (
            row[concatenated_column - 1] if concatenated_column is not None else None
        )
        # Nel modello le concatenazioni sono spesso formule. In modalità di
        # scrittura openpyxl restituisce la formula e non il valore calcolato:
        # costruiamo quindi esplicitamente il testo normalizzato.
        if isinstance(raw_concatenated, str) and raw_concatenated.startswith("="):
            concatenated = f"{code}, {display}"
        else:
            concatenated = clean_text(raw_concatenated) or f"{code}, {display}"
        item = LookupValue(code=code, display=display, concatenated=concatenated, order=order)
        order += 1
        by_code[code] = item
        by_display[text_key(display)] = item
        by_display[text_key(concatenated)] = item
    return LookupTable(by_code=by_code, by_display=by_display)


def extend_lookup_table(
    lookup: LookupTable, additions: dict[str, str], code_width: int
) -> None:
    next_order = 1 + max((item.order for item in lookup.by_code.values()), default=-1)
    for raw_code, raw_display in additions.items():
        code = code_value(raw_code, code_width)
        if code in lookup.by_code:
            continue
        display = clean_text(raw_display)
        item = LookupValue(
            code=code,
            display=display,
            concatenated=f"{code}, {display}",
            order=next_order,
        )
        next_order += 1
        lookup.by_code[code] = item
        lookup.by_display[text_key(display)] = item
        lookup.by_display[text_key(item.concatenated)] = item


def append_lookup_additions(
    worksheet: Any, additions: dict[str, str], code_width: int
) -> int:
    """Aggiunge al foglio anagrafico i codici configurati ma assenti."""
    current = load_lookup(worksheet, code_width=code_width)
    for raw_code, raw_display in additions.items():
        code = code_value(raw_code, code_width)
        if code in current.by_code:
            continue
        target_row = worksheet.max_row + 1
        style_row = max(3, target_row - 1)
        for column in range(1, worksheet.max_column + 1):
            source_cell = worksheet.cell(style_row, column)
            target_cell = worksheet.cell(target_row, column)
            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)
        display = clean_text(raw_display)
        worksheet.cell(target_row, 2).value = code
        worksheet.cell(target_row, 3).value = display
        worksheet.cell(target_row, 4).value = f"{code}, {display}"
        item = LookupValue(
            code=code,
            display=display,
            concatenated=f"{code}, {display}",
            order=len(current.by_code),
        )
        current.by_code[code] = item
    return worksheet.max_row


@dataclass
class IssueBook:
    counts: dict[str, Counter[str]]
    rows: dict[str, dict[str, list[int]]]

    @classmethod
    def create(cls) -> "IssueBook":
        return cls(counts=defaultdict(Counter), rows=defaultdict(lambda: defaultdict(list)))

    def add(self, category: str, value: Any, row_number: int) -> None:
        shown_value = clean_text(value) or "<vuoto>"
        self.counts[category][shown_value] += 1
        if len(self.rows[category][shown_value]) < 8:
            self.rows[category][shown_value].append(row_number)

    def has_issues(self) -> bool:
        return any(counter for counter in self.counts.values())

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for category, counter in self.counts.items():
            result[category] = [
                {
                    "value": value,
                    "cases": cases,
                    "sample_source_rows": self.rows[category][value],
                }
                for value, cases in counter.most_common()
            ]
        return result


@dataclass(frozen=True)
class AggregateKey:
    assistance_asl_code: str
    provider_asl_code: str
    age_code: str
    gender_code: str
    structure_code: str


@dataclass
class ProcessingResult:
    total_cases: int
    mapped_cases: int
    groups: Counter[AggregateKey]
    issues: IssueBook


def process_source(source_path: Path, template_path: Path, cfg: dict[str, Any]) -> ProcessingResult:
    validation = validate_source_workbook(source_path, cfg)
    if not validation.valid:
        raise SourceValidationError(validation, cfg["source_columns"])
    assert validation.sheet_name is not None
    assert validation.header_row is not None

    source_workbook = load_workbook(source_path, read_only=True, data_only=True)
    template_workbook = load_workbook(template_path, read_only=True, data_only=True)
    try:
        source_sheet = source_workbook[validation.sheet_name]
        lookup_names = cfg["lookup_sheets"]
        asl_lookup = load_lookup(template_workbook[lookup_names["asl"]], code_width=6)
        extend_lookup_table(asl_lookup, cfg.get("supplemental_asl", {}), 6)
        age_lookup = load_lookup(template_workbook[lookup_names["age"]])
        gender_lookup = load_lookup(template_workbook[lookup_names["gender"]])
        structure_lookup = load_lookup(
            template_workbook[lookup_names["structure"]], code_width=6
        )

        header_row = validation.header_row
        columns = validation.columns
        structure_translation: dict[str, tuple[str, str]] = {}
        for source_name, translation in cfg.get("structure_translation", {}).items():
            if not isinstance(translation, dict):
                raise RuntimeError(
                    f"La traduzione di {source_name!r} deve indicare "
                    "provider_asl_code e structure_code."
                )
            structure_translation[text_key(source_name)] = (
                code_value(translation.get("provider_asl_code"), 6),
                code_value(translation.get("structure_code"), 6),
            )

        groups: Counter[AggregateKey] = Counter()
        issues = IssueBook.create()
        total_cases = 0
        mapped_cases = 0

        for row_number, row in enumerate(
            source_sheet.iter_rows(min_row=header_row + 1, values_only=True),
            start=header_row + 1,
        ):
            patient_id = row[columns["patient_id"] - 1]
            if patient_id is None or clean_text(patient_id) == "":
                continue
            total_cases += 1
            row_has_issue = False

            try:
                region_code = code_value(row[columns["region"] - 1], 3)
                local_asl_code = code_value(row[columns["asl"] - 1], 3)
                assistance_asl_code = region_code + local_asl_code
            except ValueError:
                raw_assistance_asl = (
                    f"{clean_text(row[columns['region'] - 1])}|"
                    f"{clean_text(row[columns['asl'] - 1])}"
                )
                issues.add("assistance_asl", raw_assistance_asl, row_number)
                assistance_asl_code = ""
                row_has_issue = True
            assistance_asl_item = asl_lookup.by_code.get(assistance_asl_code)
            if assistance_asl_code and assistance_asl_item is None:
                issues.add("assistance_asl", assistance_asl_code, row_number)
                row_has_issue = True

            raw_age = row[columns["age"] - 1]
            age_item = age_lookup.by_display.get(text_key(raw_age))
            if age_item is None:
                issues.add("age", raw_age, row_number)
                row_has_issue = True

            raw_gender = row[columns["gender"] - 1]
            gender_item = gender_lookup.by_display.get(text_key(raw_gender))
            if gender_item is None:
                issues.add("gender", raw_gender, row_number)
                row_has_issue = True

            raw_provider = row[columns["provider"] - 1]
            provider_translation = structure_translation.get(text_key(raw_provider))
            if provider_translation is None:
                issues.add("structure", raw_provider, row_number)
                row_has_issue = True
                provider_asl_item = None
                structure_item = None
            else:
                provider_asl_code, structure_code = provider_translation
                provider_asl_item = asl_lookup.by_code.get(provider_asl_code)
                structure_item = structure_lookup.by_code.get(structure_code)
                if provider_asl_item is None:
                    issues.add("provider_asl", provider_asl_code, row_number)
                    row_has_issue = True
                if structure_item is None:
                    issues.add("structure_code", structure_code, row_number)
                    row_has_issue = True

            if row_has_issue:
                continue

            assert (
                assistance_asl_item
                and provider_asl_item
                and age_item
                and gender_item
                and structure_item
            )
            groups[
                AggregateKey(
                    assistance_asl_code=assistance_asl_item.code,
                    provider_asl_code=provider_asl_item.code,
                    age_code=age_item.code,
                    gender_code=gender_item.code,
                    structure_code=structure_item.code,
                )
            ] += 1
            mapped_cases += 1

        return ProcessingResult(
            total_cases=total_cases,
            mapped_cases=mapped_cases,
            groups=groups,
            issues=issues,
        )
    finally:
        source_workbook.close()
        template_workbook.close()


def print_issues(issues: IssueBook) -> None:
    labels = {
        "assistance_asl": "ASL di assistenza non riconosciute",
        "age": "Fasce d'età non riconosciute",
        "gender": "Generi non riconosciuti",
        "structure": "Strutture/eroganti non riconosciuti",
        "provider_asl": "ASL eroganti configurate ma assenti dall'anagrafica",
        "structure_code": "Codici struttura configurati ma assenti dall'anagrafica",
    }
    for category, counter in issues.counts.items():
        if not counter:
            continue
        print(f"\n{labels.get(category, category)}:")
        for value, cases in counter.most_common():
            rows = ", ".join(str(number) for number in issues.rows[category][value])
            print(f"  - {value}: {cases} casi (righe sorgente: {rows})")


def clone_template_cell(source_cell: Any, target_cell: Any, target_row: int) -> None:
    value = source_cell.value
    if isinstance(value, str) and value.startswith("=") and target_row != source_cell.row:
        value = Translator(value, origin=source_cell.coordinate).translate_formula(
            target_cell.coordinate
        )
    target_cell.value = value
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)
    target_cell.number_format = source_cell.number_format
    target_cell.font = copy(source_cell.font)
    target_cell.fill = copy(source_cell.fill)
    target_cell.border = copy(source_cell.border)
    target_cell.alignment = copy(source_cell.alignment)
    target_cell.protection = copy(source_cell.protection)


def extend_data_validations(worksheet: Any, last_row: int) -> None:
    if worksheet.data_validations is None:
        return
    for validation in worksheet.data_validations.dataValidation:
        updated_ranges: list[str] = []
        for cell_range in validation.ranges.ranges:
            current = CellRange(str(cell_range))
            if current.min_row <= 2 <= current.max_row and current.min_col == current.max_col:
                current.max_row = max(current.max_row, last_row)
            updated_ranges.append(str(current))
        validation.sqref = MultiCellRange(" ".join(updated_ranges))


def restore_worksheet_extensions(
    template_path: Path,
    output_path: Path,
    lookup_last_rows: dict[str, int] | None = None,
) -> None:
    """Ripristina le estensioni XML non supportate da openpyxl.

    Il modello PNT usa x14:dataValidations per i menu a tendina. Poiché lo
    script modifica solo i dati e non la struttura delle convalide, le sezioni
    extLst originali possono essere reinserite senza trasformazioni.
    """
    extensions: dict[str, bytes] = {}
    with zipfile.ZipFile(template_path, "r") as template_zip:
        for name in template_zip.namelist():
            if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                continue
            match = WORKSHEET_EXTENSION_RE.search(template_zip.read(name))
            if match:
                extension = match.group(0)
                opening_tag_end = extension.find(b">")
                opening_tag = extension[:opening_tag_end]
                if b"xmlns:xr=" not in opening_tag:
                    extension = extension.replace(
                        b"<extLst",
                        b'<extLst xmlns:xr="' + XR_NAMESPACE + b'"',
                        1,
                    )
                for sheet_name, last_row in (lookup_last_rows or {}).items():
                    range_pattern = re.compile(
                        re.escape(sheet_name.encode("utf-8"))
                        + rb"!\$D\$3:\$D\$\d+"
                    )
                    extension = range_pattern.sub(
                        sheet_name.encode("utf-8")
                        + f"!$D$3:$D${last_row}".encode("utf-8"),
                        extension,
                    )
                extensions[name] = extension

    if not extensions:
        return

    output_path.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=output_path.stem + "_", suffix=".xlsx", dir=output_path.parent
    )
    os.close(file_descriptor)
    temporary_path = Path(temporary_name)
    try:
        with zipfile.ZipFile(output_path, "r") as source_zip, zipfile.ZipFile(
            temporary_path, "w"
        ) as destination_zip:
            for info in source_zip.infolist():
                data = source_zip.read(info.filename)
                extension = extensions.get(info.filename)
                if extension:
                    if WORKSHEET_EXTENSION_RE.search(data):
                        data = WORKSHEET_EXTENSION_RE.sub(extension, data, count=1)
                    else:
                        data = data.replace(
                            b"</worksheet>", extension + b"</worksheet>", 1
                        )
                destination_zip.writestr(info, data)
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def write_output(
    template_path: Path,
    output_path: Path,
    cfg: dict[str, Any],
    result: ProcessingResult,
) -> tuple[int, int]:
    workbook = load_workbook(template_path, data_only=False)
    try:
        worksheet = workbook[cfg["target_sheet"]]
        lookup_names = cfg["lookup_sheets"]
        asl_worksheet = workbook[lookup_names["asl"]]
        asl_lookup_last_row = append_lookup_additions(
            asl_worksheet, cfg.get("supplemental_asl", {}), 6
        )
        asl_lookup = load_lookup(asl_worksheet, code_width=6)
        age_lookup = load_lookup(workbook[lookup_names["age"]])
        gender_lookup = load_lookup(workbook[lookup_names["gender"]])
        structure_lookup = load_lookup(workbook[lookup_names["structure"]], code_width=6)

        template_row = 2
        old_max_row = worksheet.max_row
        max_column = worksheet.max_column
        base_cells = [copy(worksheet.cell(template_row, column)) for column in range(1, max_column + 1)]
        base_height = worksheet.row_dimensions[template_row].height

        age_order = {item.code: item.order for item in age_lookup.by_code.values()}
        gender_order = {item.code: item.order for item in gender_lookup.by_code.values()}
        structure_order = {item.code: item.order for item in structure_lookup.by_code.values()}
        asl_order = {item.code: item.order for item in asl_lookup.by_code.values()}
        sorted_groups = sorted(
            result.groups.items(),
            key=lambda pair: (
                pair[0].provider_asl_code,
                structure_order[pair[0].structure_code],
                asl_order[pair[0].assistance_asl_code],
                age_order[pair[0].age_code],
                gender_order[pair[0].gender_code],
            ),
        )

        target_columns = {
            field: column_index_from_string(letter)
            for field, letter in cfg["target_columns"].items()
        }
        for offset, (key, quantity) in enumerate(sorted_groups):
            target_row = template_row + offset
            for column, base_cell in enumerate(base_cells, start=1):
                clone_template_cell(base_cell, worksheet.cell(target_row, column), target_row)
            if base_height is not None:
                worksheet.row_dimensions[target_row].height = base_height
            worksheet.row_dimensions[target_row].hidden = False
            worksheet.cell(target_row, target_columns["age"]).value = age_lookup.by_code[
                key.age_code
            ].concatenated
            normalized_provider_asl = asl_lookup.by_code[
                key.provider_asl_code
            ].concatenated
            worksheet.cell(target_row, target_columns["assistance_asl"]).value = (
                asl_lookup.by_code[key.assistance_asl_code].concatenated
            )
            worksheet.cell(target_row, target_columns["gender"]).value = gender_lookup.by_code[
                key.gender_code
            ].concatenated
            worksheet.cell(target_row, target_columns["quantity"]).value = quantity
            worksheet.cell(target_row, target_columns["provider_asl"]).value = (
                normalized_provider_asl
            )
            worksheet.cell(target_row, target_columns["structure"]).value = structure_lookup.by_code[
                key.structure_code
            ].concatenated

        last_row = template_row + len(sorted_groups) - 1
        worksheet.auto_filter.ref = (
            f"A1:{get_column_letter(max_column)}{last_row}"
        )
        worksheet.auto_filter.filterColumn = []
        worksheet.auto_filter.sortState = None
        if last_row < old_max_row:
            for row in worksheet.iter_rows(
                min_row=last_row + 1, max_row=old_max_row, min_col=1, max_col=max_column
            ):
                for cell in row:
                    cell.value = None

        extend_data_validations(worksheet, last_row)
        for table in worksheet.tables.values():
            table_range = CellRange(table.ref)
            if table_range.min_row == 1:
                table_range.max_row = last_row
                table.ref = str(table_range)

        temporary_handle = tempfile.NamedTemporaryFile(
            prefix="tracciato_telemedicina_", suffix=".xlsx", delete=False
        )
        temporary_workbook_path = Path(temporary_handle.name)
        temporary_handle.close()
        workbook.save(temporary_workbook_path)
    finally:
        workbook.close()

    try:
        restore_worksheet_extensions(
            template_path,
            temporary_workbook_path,
            {lookup_names["asl"]: asl_lookup_last_row},
        )

        verification_workbook = load_workbook(
            temporary_workbook_path, read_only=False, data_only=True
        )
        try:
            verification_sheet = verification_workbook[cfg["target_sheet"]]
            quantity_column = column_index_from_string(cfg["target_columns"]["quantity"])
            quantity_total = sum(
                int(verification_sheet.cell(row, quantity_column).value or 0)
                for row in range(2, 2 + len(result.groups))
            )
            visible_quantity_total = sum(
                int(verification_sheet.cell(row, quantity_column).value or 0)
                for row in range(2, 2 + len(result.groups))
                if not verification_sheet.row_dimensions[row].hidden
            )
            if visible_quantity_total != quantity_total:
                raise RuntimeError(
                    "Verifica output fallita: il totale delle righe visibili "
                    f"è {visible_quantity_total}, mentre il totale complessivo "
                    f"è {quantity_total}."
                )
        finally:
            verification_workbook.close()

        # I file OneDrive possono andare in timeout se openpyxl costruisce lo
        # ZIP direttamente sulla cartella sincronizzata. Il file viene quindi
        # prodotto e verificato in locale e pubblicato con una singola copia.
        output_path.parent.mkdir(parents=True, exist_ok=True)
        staging_path = output_path.with_name(
            f".{output_path.name}.{os.getpid()}.staging"
        )
        try:
            shutil.copy2(temporary_workbook_path, staging_path)
            os.replace(staging_path, output_path)
        finally:
            if staging_path.exists():
                staging_path.unlink()
        return len(result.groups), quantity_total
    finally:
        if temporary_workbook_path.exists():
            temporary_workbook_path.unlink()


def write_report(
    report_path: Path,
    source_path: Path,
    template_path: Path,
    result: ProcessingResult,
) -> None:
    report = {
        "source_file": str(source_path),
        "template_file": str(template_path),
        "total_cases": result.total_cases,
        "mapped_cases": result.mapped_cases,
        "unmapped_cases": result.total_cases - result.mapped_cases,
        "aggregated_tuples": len(result.groups),
        "issues": result.issues.to_dict(),
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def package_resource(name: str) -> Path:
    """Restituisce una risorsa sia da sorgente sia da bundle PyInstaller."""
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        return Path(bundle_root) / "telemedicina_verticale" / name
    return Path(__file__).resolve().parent / name


def load_default_config() -> dict[str, Any]:
    return load_vertical_config(package_resource("config.json"))


def load_vertical_config(path: str | Path) -> dict[str, Any]:
    config = read_json(Path(path).expanduser().resolve())
    try:
        return config["verticals"]["teleconsulto_tat"]
    except KeyError as exc:
        raise RuntimeError(
            "Configurazione Telemedicina non valida: manca verticals.teleconsulto_tat."
        ) from exc


def format_issues(issues: IssueBook) -> str:
    labels = {
        "assistance_asl": "ASL di assistenza non riconosciute",
        "age": "Fasce d'età non riconosciute",
        "gender": "Generi non riconosciuti",
        "structure": "Eroganti/strutture non configurati",
        "provider_asl": "ASL eroganti configurate ma assenti dall'anagrafica",
        "structure_code": "Codici struttura configurati ma assenti dall'anagrafica",
    }
    lines: list[str] = []
    for category, counter in issues.counts.items():
        if not counter:
            continue
        lines.append(labels.get(category, category) + ":")
        for value, cases in counter.most_common():
            rows = ", ".join(str(number) for number in issues.rows[category][value])
            lines.append(f"- {value}: {cases} casi (righe sorgente: {rows})")
        lines.append("")
    return "\n".join(lines).rstrip()


@dataclass(frozen=True)
class GenerationSummary:
    source_path: Path
    source_sheet: str
    header_row: int
    output_path: Path
    total_cases: int
    mapped_cases: int
    aggregated_rows: int
    output_quantity_total: int


def generate_trace(
    source_path: str | Path,
    output_path: str | Path,
    *,
    template_path: str | Path | None = None,
    cfg: dict[str, Any] | None = None,
    log: Any | None = None,
) -> GenerationSummary:
    """Valida, aggrega e genera il tracciato PNT dal file scelto in GUI."""
    active_cfg = cfg or load_default_config()
    source = Path(source_path).expanduser().resolve()
    output = Path(output_path).expanduser().resolve()
    template = (
        Path(template_path).expanduser().resolve()
        if template_path is not None
        else package_resource("Tracciato_PNT_template.xlsx")
    )

    if not template.is_file():
        raise RuntimeError(f"Modello PNT non trovato: {template}")
    if output.suffix.casefold() != ".xlsx":
        raise RuntimeError("Il file di output deve avere estensione .xlsx.")
    if output == source or output == template.resolve():
        raise RuntimeError("Il file di output deve essere diverso dalla sorgente e dal modello.")

    validation = validate_source_workbook(source, active_cfg)
    if not validation.valid:
        raise SourceValidationError(validation, active_cfg["source_columns"])
    assert validation.sheet_name is not None
    assert validation.header_row is not None

    emit = log or (lambda _message: None)
    emit(f"Sorgente conforme: {source}")
    emit(f"Foglio dati: {validation.sheet_name}; intestazioni alla riga {validation.header_row}")
    emit("Lettura per nome delle sole colonne utilizzate; colonne aggiuntive ignorate.")

    result = process_source(source, template, active_cfg)
    emit(f"Casi letti (ID paziente valorizzato): {result.total_cases}")
    emit(f"Casi ricondotti: {result.mapped_cases}")
    if result.total_cases == 0:
        raise RuntimeError("La sorgente non contiene casi con ID paziente valorizzato.")
    if result.issues.has_issues():
        raise MappingError(
            "Alcuni valori non sono riconducibili. Aggiornare config.json nelle traduzioni "
            "e ripetere l'elaborazione.\n\n"
            + format_issues(result.issues)
        )
    if sum(result.groups.values()) != result.total_cases:
        raise RuntimeError(
            "Controllo interno fallito: la somma delle tuple non coincide con i casi letti."
        )

    aggregated_rows, quantity_total = write_output(template, output, active_cfg, result)
    if quantity_total != result.total_cases:
        raise RuntimeError(
            "Verifica output fallita: il totale della colonna H "
            f"({quantity_total}) non coincide con i casi sorgente ({result.total_cases})."
        )
    emit(f"Righe aggregate prodotte: {aggregated_rows}")
    emit(f"Totale colonna H verificato: {quantity_total}")
    emit(f"Output: {output}")
    return GenerationSummary(
        source_path=source,
        source_sheet=validation.sheet_name,
        header_row=validation.header_row,
        output_path=output,
        total_cases=result.total_cases,
        mapped_cases=result.mapped_cases,
        aggregated_rows=aggregated_rows,
        output_quantity_total=quantity_total,
    )


def parse_args() -> argparse.Namespace:
    default_config = package_resource("config.json")
    parser = argparse.ArgumentParser(
        description="Aggrega il foglio sorgente e compila il tracciato PNT."
    )
    parser.add_argument("--input-dir", type=Path, default=Path.cwd())
    parser.add_argument("--config", type=Path, default=default_config)
    parser.add_argument("--vertical", default="teleconsulto_tat")
    parser.add_argument("--source", type=Path, help="File sorgente; altrimenti viene rilevato dai fogli.")
    parser.add_argument("--template", type=Path, help="File del tracciato; altrimenti viene rilevato dai fogli.")
    parser.add_argument("--output", type=Path, help="Percorso del nuovo file compilato.")
    parser.add_argument("--report", type=Path, help="Report JSON di controllo/mappatura.")
    parser.add_argument("--audit", action="store_true", help="Controlla le mappature senza scrivere il tracciato.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = read_json(args.config.resolve())
    try:
        cfg = config["verticals"][args.vertical]
    except KeyError as exc:
        available = ", ".join(sorted(config.get("verticals", {})))
        raise RuntimeError(
            f"Verticale {args.vertical!r} non configurato. Disponibili: {available}"
        ) from exc

    input_dir = args.input_dir.resolve()
    source_path = discover_workbook(
        input_dir,
        {cfg["source_sheet"]},
        args.source,
        "sorgente",
    )
    template_required = {cfg["target_sheet"], *cfg["lookup_sheets"].values()}
    template_path = discover_workbook(
        input_dir,
        template_required,
        args.template,
        "tracciato",
    )

    print(f"Sorgente:  {source_path}")
    print(f"Tracciato: {template_path}")
    result = process_source(source_path, template_path, cfg)
    print(f"Casi letti (ID paziente non vuoto): {result.total_cases}")
    print(f"Casi ricondotti:                    {result.mapped_cases}")
    print(f"Tuple aggregate ricondotte:         {len(result.groups)}")

    expected = cfg.get("expected_case_count")
    if expected is not None and result.total_cases != int(expected):
        raise RuntimeError(
            f"Totale sorgente inatteso: {result.total_cases}; configurato: {expected}."
        )

    if args.report:
        write_report(args.report.resolve(), source_path, template_path, result)
        print(f"Report: {args.report.resolve()}")

    if result.issues.has_issues():
        print_issues(result.issues)
        print(
            "\nElaborazione sospesa: completare le traduzioni nel file di configurazione.",
            file=sys.stderr,
        )
        return 2

    if sum(result.groups.values()) != result.total_cases:
        raise RuntimeError("Controllo interno fallito: la somma delle tuple non coincide con i casi.")

    if args.audit:
        print("Audit completato: tutte le righe sono riconducibili.")
        return 0

    output_path = args.output
    if output_path is None:
        output_path = (
            input_dir
            / "outputs"
            / f"{template_path.stem}_compilato_{args.vertical}.xlsx"
        )
    output_path = output_path.resolve()
    group_count, quantity_total = write_output(
        template_path, output_path, cfg, result
    )
    if quantity_total != result.total_cases:
        raise RuntimeError(
            f"Verifica output fallita: somma H={quantity_total}, casi={result.total_cases}."
        )
    print(f"Righe aggregate scritte: {group_count}")
    print(f"Totale colonna H:        {quantity_total}")
    print(f"Output:                  {output_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as error:
        print(f"ERRORE: {error}", file=sys.stderr)
        raise SystemExit(1)
