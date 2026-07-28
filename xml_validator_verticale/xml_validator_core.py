from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from typing import Callable
import xml.etree.ElementTree as ET

try:
    import xmlschema
except ModuleNotFoundError:  # Consente di mostrare un errore esplicito dalla GUI.
    xmlschema = None


XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"
XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance"


@dataclass(frozen=True)
class ValidationIssue:
    path: str
    message: str
    line: int | None = None

    def display_text(self, number: int) -> str:
        parts = [f"[{number}]"]
        if self.line is not None:
            parts.append(f"riga {self.line}")
        if self.path:
            parts.append(self.path)
        parts.append(self.message)
        return " | ".join(parts)


@dataclass(frozen=True)
class ValidationResult:
    issues: tuple[ValidationIssue, ...]
    source_namespaces: tuple[str, ...]
    schema_namespace: str | None

    @property
    def namespace_was_present(self) -> bool:
        return bool(self.source_namespaces)


def _split_expanded_name(name: str) -> tuple[str | None, str]:
    if name.startswith("{") and "}" in name:
        namespace, local = name[1:].split("}", 1)
        return namespace, local
    if ":" in name:
        _prefix, local = name.split(":", 1)
        return None, local
    return None, name


def _expanded_name(namespace: str | None, local: str) -> str:
    return f"{{{namespace}}}{local}" if namespace else local


def _schema_namespace_settings(xsd_path: Path) -> tuple[str | None, bool, bool]:
    root = ET.parse(xsd_path).getroot()
    target_namespace = root.attrib.get("targetNamespace") or None
    elements_qualified = root.attrib.get("elementFormDefault", "unqualified") == "qualified"
    attributes_qualified = root.attrib.get("attributeFormDefault", "unqualified") == "qualified"
    return target_namespace, elements_qualified, attributes_qualified


def _normalise_element_namespaces(
    element: ET.Element,
    *,
    target_namespace: str | None,
    elements_qualified: bool,
    attributes_qualified: bool,
    depth: int = 0,
) -> set[str]:
    source_namespaces: set[str] = set()

    if isinstance(element.tag, str):
        source_namespace, local = _split_expanded_name(element.tag)
        if source_namespace and source_namespace not in {XML_NAMESPACE, XSI_NAMESPACE}:
            source_namespaces.add(source_namespace)

        # Un elemento globale (la radice) appartiene sempre al targetNamespace.
        # Gli elementi locali seguono elementFormDefault dello schema.
        output_namespace = target_namespace if target_namespace and (depth == 0 or elements_qualified) else None
        element.tag = _expanded_name(output_namespace, local)

    normalised_attributes: dict[str, str] = {}
    for raw_name, value in element.attrib.items():
        source_namespace, local = _split_expanded_name(raw_name)
        if source_namespace and source_namespace not in {XML_NAMESPACE, XSI_NAMESPACE}:
            source_namespaces.add(source_namespace)

        if source_namespace in {XML_NAMESPACE, XSI_NAMESPACE}:
            output_name = _expanded_name(source_namespace, local)
        else:
            output_namespace = target_namespace if target_namespace and attributes_qualified else None
            output_name = _expanded_name(output_namespace, local)

        if output_name in normalised_attributes:
            raise ValueError(
                f"Attributi omonimi dopo la rimozione del namespace: '{local}' "
                f"nell'elemento '{element.tag}'."
            )
        normalised_attributes[output_name] = value

    element.attrib.clear()
    element.attrib.update(normalised_attributes)

    for child in element:
        source_namespaces.update(
            _normalise_element_namespaces(
                child,
                target_namespace=target_namespace,
                elements_qualified=elements_qualified,
                attributes_qualified=attributes_qualified,
                depth=depth + 1,
            )
        )

    return source_namespaces


def prepare_namespace_neutral_xml(xml_path: Path, xsd_path: Path, output_path: Path) -> tuple[tuple[str, ...], str | None]:
    """Crea una copia temporanea dell'XML ignorando i namespace dichiarati in input.

    Se lo schema possiede un targetNamespace, gli elementi vengono riallineati a
    quel namespace. In caso contrario vengono scritti senza namespace. Il file
    sorgente non viene mai modificato.
    """

    target_namespace, elements_qualified, attributes_qualified = _schema_namespace_settings(xsd_path)
    tree = ET.parse(xml_path)
    source_namespaces = _normalise_element_namespaces(
        tree.getroot(),
        target_namespace=target_namespace,
        elements_qualified=elements_qualified,
        attributes_qualified=attributes_qualified,
    )
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    return tuple(sorted(source_namespaces)), target_namespace


def _clean_validation_path(path: str) -> str:
    path = re.sub(r"\{[^}]+\}", "", path)
    path = re.sub(r"(?<=/)[A-Za-z_][\w.-]*:", "", path)
    path = re.sub(r"(?<=@)[A-Za-z_][\w.-]*:", "", path)
    return path


def _clean_validation_message(message: str) -> str:
    return re.sub(r"\{[^}]+\}", "", message)


def _validation_issue(error: object) -> ValidationIssue:
    path = _clean_validation_path(str(getattr(error, "path", "") or ""))
    reason = getattr(error, "reason", None)
    message = _clean_validation_message(str(reason or error))
    line = getattr(error, "sourceline", None)
    if not isinstance(line, int):
        line = None
    return ValidationIssue(path=path, message=message, line=line)


def validate_xml_without_namespaces(
    xml_path: str | Path,
    xsd_path: str | Path,
    *,
    on_issue: Callable[[ValidationIssue], None] | None = None,
) -> ValidationResult:
    """Valida un XML contro uno XSD senza considerare i namespace dell'XML.

    La normalizzazione avviene su una copia temporanea. ``on_issue`` viene
    invocata man mano che xmlschema produce gli errori, così una GUI può
    visualizzarli senza attendere la fine della validazione.
    """

    if xmlschema is None:
        raise RuntimeError(
            "La validazione richiede il pacchetto 'xmlschema'. "
            "Installa le dipendenze con: python -m pip install -r requirements.txt"
        )

    xml_file = Path(xml_path).expanduser()
    xsd_file = Path(xsd_path).expanduser()
    if not xsd_file.is_file():
        raise FileNotFoundError(f"File XSD non trovato: {xsd_file}")
    if not xml_file.is_file():
        raise FileNotFoundError(f"File XML non trovato: {xml_file}")

    with TemporaryDirectory(prefix="xml-validator-") as temp_dir:
        normalised_xml = Path(temp_dir) / "namespace_neutral.xml"
        source_namespaces, schema_namespace = prepare_namespace_neutral_xml(
            xml_file,
            xsd_file,
            normalised_xml,
        )
        schema = xmlschema.XMLSchema(xsd_file)

        issues: list[ValidationIssue] = []
        for error in schema.iter_errors(normalised_xml):
            issue = _validation_issue(error)
            issues.append(issue)
            if on_issue is not None:
                on_issue(issue)

    return ValidationResult(
        issues=tuple(issues),
        source_namespaces=source_namespaces,
        schema_namespace=schema_namespace,
    )
