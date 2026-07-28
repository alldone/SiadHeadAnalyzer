from __future__ import annotations

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from siad_report_gui import (
    XmlFileInfo,
    age_on_reference_date,
    build_report,
    expected_cf_control_char,
    extract_cf_from_id_rec,
    infer_birth_year_from_cf,
    quarter_from_path,
    scan_xml_files,
    validate_xml_files,
)


class SiadReportTests(unittest.TestCase):
    def _write_track1_xml(self, path: Path, cf: str, *, namespace: str = "") -> None:
        xmlns = f' xmlns="{namespace}"' if namespace else ""
        path.write_text(
            f"""<?xml version="1.0" encoding="UTF-8"?>
<FlsAssDom_1{xmlns}>
  <Assistenza>
    <Assistito><DatiAnagrafici><AnnoNascita>1980</AnnoNascita></DatiAnagrafici></Assistito>
    <Erogatore><CodiceASL>201</CodiceASL></Erogatore>
    <Eventi><PresaInCarico data="2026-01-10"><Id_Rec>18020120260110{cf}</Id_Rec></PresaInCarico></Eventi>
  </Assistenza>
</FlsAssDom_1>
""",
            encoding="utf-8",
        )

    def test_scan_accepts_uppercase_xml_extension(self) -> None:
        with TemporaryDirectory() as temp_dir:
            base = Path(temp_dir)
            xml_path = base / "SIAD_TEST.XML"
            self._write_track1_xml(xml_path, "RSSMRA80A01H501U")

            files = scan_xml_files(base, {"FlsAssDom_1": 1})

            self.assertEqual(1, len(files))
            self.assertEqual(1, files[0].track)

    def test_quarter_is_read_from_siad_filename(self) -> None:
        self.assertEqual("T2", quarter_from_path("cartella/SIAD20326T21_SIAD_1_test.XML"))
        self.assertEqual("T1", quarter_from_path("cartella/SIAD20326T12_SIAD_2_test.xml"))

    def test_omocodia_is_accepted_and_decoded_for_age(self) -> None:
        cf_without_control = "RSSMRAU0A01H501"
        control_char = expected_cf_control_char(cf_without_control + "A")
        self.assertIsNotNone(control_char)
        cf_with_birth_year_omocodia = cf_without_control + str(control_char)

        self.assertEqual(cf_with_birth_year_omocodia, extract_cf_from_id_rec(cf_with_birth_year_omocodia))
        self.assertEqual(80, infer_birth_year_from_cf(cf_with_birth_year_omocodia))
        self.assertEqual(46, age_on_reference_date(cf_with_birth_year_omocodia, 1980, date(2026, 12, 31)))

    def test_missing_clinical_tags_do_not_exclude_assisted_person(self) -> None:
        with TemporaryDirectory() as temp_dir:
            xml_path = Path(temp_dir) / "SIAD20126T11_SIAD_1_test.xml"
            self._write_track1_xml(xml_path, "RSSMRA80A01H501U")
            file_info = XmlFileInfo(xml_path, xml_path.name, 1, "FlsAssDom_1")

            summary, details, _unique, _global_heads, extraction_issues = build_report([file_info], 2026)

            self.assertEqual(1, len(details))
            self.assertEqual([], extraction_issues)
            self.assertEqual(1, summary[0]["TOT. PRESE IN CARICO attive nel 2026"])
            self.assertEqual(1, summary[0]["[CF per azienda] TOT. PAZIENTI* attivi nel 2026"])

    def test_unrecognised_cf_is_reported_as_extraction_issue(self) -> None:
        with TemporaryDirectory() as temp_dir:
            xml_path = Path(temp_dir) / "SIAD20126T11_SIAD_1_test.xml"
            self._write_track1_xml(xml_path, "RSSMRA80A01H5011")
            file_info = XmlFileInfo(xml_path, xml_path.name, 1, "FlsAssDom_1")

            summary, details, _unique, _global_heads, extraction_issues = build_report([file_info], 2026)

            self.assertEqual([], summary)
            self.assertEqual([], details)
            self.assertEqual(1, len(extraction_issues))
            self.assertEqual(1, extraction_issues[0].record_number)
            self.assertIn("non riconosciuto", extraction_issues[0].reason)

    def test_xsd_errors_are_returned_separately_from_calculation(self) -> None:
        with TemporaryDirectory() as temp_dir:
            base = Path(temp_dir)
            xml_path = base / "SIAD_TEST.xml"
            xsd_path = base / "schema.xsd"
            self._write_track1_xml(xml_path, "RSSMRA80A01H501U")
            xsd_path.write_text(
                """<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element name="FlsAssDom_1">
    <xs:complexType><xs:sequence><xs:element name="TagClinicoObbligatorio"/></xs:sequence></xs:complexType>
  </xs:element>
</xs:schema>
""",
                encoding="utf-8",
            )
            file_info = XmlFileInfo(xml_path, xml_path.name, 1, "FlsAssDom_1")

            validation = validate_xml_files([file_info], {1: xsd_path})
            summary, _details, _unique, _global_heads, extraction_issues = build_report([file_info], 2026)

            self.assertEqual("NON VALIDO", validation[0].status)
            self.assertTrue(validation[0].errors)
            self.assertEqual([], extraction_issues)
            self.assertEqual(1, summary[0]["[CF per azienda] TOT. PAZIENTI* attivi nel 2026"])


if __name__ == "__main__":
    unittest.main()
