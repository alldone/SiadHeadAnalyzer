from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from xml_validator_verticale.xml_validator_core import validate_xml_without_namespaces


NO_NAMESPACE_XSD = """\
<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element name="documento">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="numero" type="xs:integer"/>
      </xs:sequence>
      <xs:attribute name="codice" type="xs:string" use="required"/>
    </xs:complexType>
  </xs:element>
</xs:schema>
"""


TARGET_NAMESPACE_XSD = """\
<?xml version="1.0" encoding="UTF-8"?>
<xs:schema
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:t="urn:schema-atteso"
    targetNamespace="urn:schema-atteso"
    elementFormDefault="qualified">
  <xs:element name="documento">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="numero" type="xs:integer"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>
"""


class XmlValidatorCoreTests(unittest.TestCase):
    def _validate(self, xsd_text: str, xml_text: str, callback=None):
        with TemporaryDirectory() as temp_dir:
            base = Path(temp_dir)
            xsd_path = base / "schema.xsd"
            xml_path = base / "documento.xml"
            xsd_path.write_text(xsd_text, encoding="utf-8")
            xml_path.write_text(xml_text, encoding="utf-8")
            original_xml = xml_path.read_bytes()

            result = validate_xml_without_namespaces(xml_path, xsd_path, on_issue=callback)

            self.assertEqual(original_xml, xml_path.read_bytes(), "Il file XML sorgente è stato modificato")
            return result

    def test_xml_without_namespace_is_valid(self) -> None:
        result = self._validate(
            NO_NAMESPACE_XSD,
            '<documento codice="A"><numero>42</numero></documento>',
        )

        self.assertEqual((), result.issues)
        self.assertFalse(result.namespace_was_present)

    def test_default_namespace_is_ignored_and_real_errors_are_reported(self) -> None:
        result = self._validate(
            NO_NAMESPACE_XSD,
            """\
<?xml version="1.0"?>
<!-- la dichiarazione del namespace non deve essere per forza sulla prima riga -->
<documento xmlns="urn:namespace-da-ignorare">
  <numero>non-un-intero</numero>
</documento>
""",
        )

        self.assertTrue(result.namespace_was_present)
        self.assertEqual(("urn:namespace-da-ignorare",), result.source_namespaces)
        self.assertGreaterEqual(len(result.issues), 2)
        messages = "\n".join(issue.message for issue in result.issues)
        self.assertRegex(messages, r"int(?:eger)?")
        self.assertIn("codice", messages)
        self.assertNotIn("not an element of the schema", messages)

    def test_foreign_namespace_is_aligned_to_schema_target_namespace(self) -> None:
        result = self._validate(
            TARGET_NAMESPACE_XSD,
            '<documento xmlns="urn:namespace-diverso"><numero>errore</numero></documento>',
        )

        self.assertEqual("urn:schema-atteso", result.schema_namespace)
        self.assertEqual(1, len(result.issues))
        self.assertRegex(result.issues[0].message, r"int(?:eger)?")
        self.assertNotIn("urn:", result.issues[0].path)

    def test_callback_receives_every_issue(self) -> None:
        streamed = []
        result = self._validate(
            NO_NAMESPACE_XSD,
            '<documento xmlns="urn:test"><numero>errore</numero></documento>',
            streamed.append,
        )

        self.assertEqual(list(result.issues), streamed)


if __name__ == "__main__":
    unittest.main()
