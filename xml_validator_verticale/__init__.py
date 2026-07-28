"""Verticale per la validazione generica di file XML rispetto a uno XSD."""

from .xml_validator_core import ValidationIssue, ValidationResult, validate_xml_without_namespaces

__all__ = ["ValidationIssue", "ValidationResult", "validate_xml_without_namespaces"]
