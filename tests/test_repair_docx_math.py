from __future__ import annotations

import importlib.util
import json
import re
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skills" / "word-math-doctor" / "scripts" / "repair_docx_math.py"
SPEC = importlib.util.spec_from_file_location("repair_docx_math", SCRIPT)
assert SPEC and SPEC.loader
repair_docx_math = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(repair_docx_math)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W, "m": M}


def make_docx(path: Path) -> None:
    document = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" mc:Ignorable="w14 wp14"><w:body>
  <w:p><w:r><w:t>勾股关系：x² + y² = z²。</w:t></w:r></w:p>
  <w:p><w:r><w:t>∫₀¹ x² dx = 1/3                                      (1)</w:t></w:r></w:p>
  <w:p><w:r><w:t>化学式 H₂O 和 CO₂ 保持为可编辑文本。</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1800" w:right="1800" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''.encode("utf-8")
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        archive.writestr("word/document.xml", document)
        archive.writestr("customXml/item1.xml", b"<fixture>unchanged-by-repair</fixture>")


def plan() -> dict:
    return {
        "schema": "word-math-doctor/repair-plan/v1",
        "operations": [
            {
                "paragraph": 1,
                "action": "inline",
                "expected_text": "勾股关系：x² + y² = z²。",
                "prefix": "勾股关系：",
                "source_formula": "x² + y² = z²",
                "suffix": "。",
                "formula": {
                    "kind": "sequence",
                    "items": [
                        {"kind": "sup", "base": "x", "exponent": "2"},
                        {"kind": "text", "value": " + "},
                        {"kind": "sup", "base": "y", "exponent": "2"},
                        {"kind": "text", "value": " = "},
                        {"kind": "sup", "base": "z", "exponent": "2"},
                    ],
                },
            },
            {
                "paragraph": 2,
                "action": "display",
                "expected_text": "∫₀¹ x² dx = 1/3                                      (1)",
                "source_formula": "∫₀¹ x² dx = 1/3",
                "number": "(1)",
                "formula": {
                    "kind": "sequence",
                    "items": [
                        {
                            "kind": "integral",
                            "lower": "0",
                            "upper": "1",
                            "body": [
                                {"kind": "sup", "base": "x", "exponent": "2"},
                                {"kind": "text", "value": " dx = "},
                                {"kind": "fraction", "numerator": "1", "denominator": "3"},
                            ],
                        }
                    ],
                },
            },
        ],
    }


class RepairDocxMathTests(unittest.TestCase):
    def run_fixture(self, mutate=lambda xml: xml, selected=None, styles=None, settings=None):
        """Run a synthetic package and return before/after bytes and the report."""
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            output = Path(directory) / "output.docx"
            plan_path = Path(directory) / "plan.json"
            make_docx(source)
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["word/document.xml"] = mutate(parts["word/document.xml"].decode()).encode()
            if styles:
                parts["word/styles.xml"] = styles.encode()
            if settings:
                parts["word/settings.xml"] = settings.encode()
            with zipfile.ZipFile(source, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            before = source.read_bytes()
            plan_path.write_text(json.dumps(selected or plan(), ensure_ascii=False), encoding="utf-8")
            try:
                report = repair_docx_math.repair(source, plan_path, output, None)
            except repair_docx_math.PlanError:
                self.assertFalse(output.exists(), "任何拒绝均不得留下修订稿")
                self.assertEqual(source.read_bytes(), before)
                raise
            with zipfile.ZipFile(output) as archive:
                return parts["word/document.xml"], archive.read("word/document.xml"), report

    def test_unselected_document_bytes_and_compatibility_root_are_untouched(self):
        before, after, report = self.run_fixture(lambda xml: xml.replace("<w:body>", "<w:body><!-- keep this comment -->"))
        old_spans = repair_docx_math.paragraph_spans(before)
        new_spans = repair_docx_math.paragraph_spans(after)
        # Compare the byte gaps around all selected paragraphs, plus unselected p3.
        self.assertEqual(before[:old_spans[0][0]], after[:new_spans[0][0]])
        self.assertEqual(before[old_spans[0][1]:old_spans[1][0]], after[new_spans[0][1]:new_spans[1][0]])
        self.assertEqual(before[old_spans[1][1]:], after[new_spans[1][1]:])
        self.assertTrue(report["validation"]["non_target_document_bytes_preserved"])

    def test_page_geometry_changes_tab_positions_without_fixed_defaults(self):
        for width, left, right in ((11906, 1440, 1440), (16838, 1000, 1200), (12240, 1800, 1800)):
            with self.subTest(width=width):
                def change(xml):
                    return xml.replace('w:w="12240"', f'w:w="{width}"').replace('w:left="1800"', f'w:left="{left}"').replace('w:right="1800"', f'w:right="{right}"')
                _, _, report = self.run_fixture(change)
                display = report["operations_applied"][1]
                self.assertEqual(display["right_tab_twips"], width-left-right)
                self.assertEqual(display["center_tab_twips"], (width-left-right)//2)

    def test_uses_target_section_not_final_section_geometry(self):
        def change(xml):
            break_paragraph = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:pPr></w:p>'
            return xml.replace('<w:p><w:r><w:t>化学式', break_paragraph + '<w:p><w:r><w:t>化学式')
        _, _, report = self.run_fixture(change)
        self.assertEqual(report["operations_applied"][1]["right_tab_twips"], 9026)

    def test_number_font_size_and_character_style_are_preserved_and_upright(self):
        def change(xml):
            return xml.replace(' (1)</w:t></w:r>', ' </w:t></w:r><w:r><w:rPr><w:rStyle w:val="EquationLabel"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/></w:rPr><w:t>(1)</w:t></w:r>')
        _, after, _ = self.run_fixture(change)
        p = list(ET.fromstring(after).iter(f"{{{W}}}p"))[1]
        label = p.findall("w:r", NS)[-1]
        self.assertEqual(label.find("w:rPr/w:rFonts", NS).get(f"{{{W}}}ascii"), "Times New Roman")
        self.assertEqual(label.find("w:rPr/w:sz", NS).get(f"{{{W}}}val"), "24")
        self.assertEqual(label.find("w:rPr/w:rStyle", NS).get(f"{{{W}}}val"), "EquationLabel")
        self.assertEqual(label.find("w:rPr/w:i", NS).get(f"{{{W}}}val"), "0")
        self.assertEqual(label.find("w:rPr/w:iCs", NS).get(f"{{{W}}}val"), "0")

    def test_inline_prefix_suffix_run_formatting_is_preserved(self):
        def change(xml):
            return xml.replace('<w:r><w:t>勾股关系：x² + y² = z²。</w:t></w:r>', '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>勾股关系：x² + y² = z²。</w:t></w:r>')
        _, after, _ = self.run_fixture(change)
        p = list(ET.fromstring(after).iter(f"{{{W}}}p"))[0]
        for run in p.findall("w:r", NS):
            self.assertIsNotNone(run.find("w:rPr/w:b", NS))
            self.assertEqual(run.find("w:rPr/w:color", NS).get(f"{{{W}}}val"), "FF0000")

    def test_unsafe_targets_and_layouts_are_rejected(self):
        mutations = {
            "bookmark": lambda x: x.replace('<w:t>勾股', '<w:fldChar w:fldCharType="begin"/><w:t>勾股'),
            "comment_anchor": lambda x: x.replace('<w:p><w:r><w:t>勾股', '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>勾股'),
            "revision": lambda x: x.replace('<w:p><w:r><w:t>勾股', '<w:p><w:r><w:rPr><w:rPrChange/></w:rPr><w:t>勾股'),
            "indent": lambda x: x.replace('<w:p><w:r><w:t>∫', '<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>∫'),
            "columns": lambda x: x.replace('<w:pgSz ', '<w:cols w:num="2"/><w:pgSz '),
            "gutter": lambda x: x.replace('w:gutter="0"', 'w:gutter="360"'),
            "missing_geometry": lambda x: re.sub(r'<w:pgMar[^>]*/>', '', x),
            "custom_tabs": lambda x: x.replace('<w:p><w:r><w:t>∫', '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="500"/></w:tabs></w:pPr><w:r><w:t>∫'),
            "table": lambda x: x.replace('<w:p><w:r><w:t>∫', '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>∫').replace('(1)</w:t></w:r></w:p>', '(1)</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
        }
        for name, change in mutations.items():
            with self.subTest(name=name), self.assertRaises(repair_docx_math.PlanError):
                self.run_fixture(change)

    def test_style_inherited_indent_is_rejected(self):
        styles = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:ind w:firstLine="420"/></w:pPr></w:style></w:styles>'
        with self.assertRaises(repair_docx_math.PlanError):
            self.run_fixture(styles=styles)

    def test_nested_section_boundary_is_rejected(self):
        def change(xml):
            nested = '<w:sdt><w:sdtContent><w:p><w:pPr><w:sectPr><w:pgSz w:w="11906"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:pPr></w:p></w:sdtContent></w:sdt>'
            return xml.replace('<w:p><w:r><w:t>化学式', nested + '<w:p><w:r><w:t>化学式')
        with self.assertRaises(repair_docx_math.PlanError):
            self.run_fixture(change)

    def test_read_or_publish_failure_leaves_no_partial_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            output = Path(directory) / "out.docx"
            plan_path = Path(directory) / "plan.json"
            make_docx(source)
            plan_path.write_text(json.dumps(plan()), encoding="utf-8")
            original_read = zipfile.ZipFile.read
            def fail_late(archive, name, *args, **kwargs):
                if name == "customXml/item1.xml":
                    raise zipfile.BadZipFile("Bad CRC-32 for file 'customXml/item1.xml'")
                return original_read(archive, name, *args, **kwargs)
            with patch.object(zipfile.ZipFile, "read", fail_late):
                with self.assertRaises(zipfile.BadZipFile):
                    repair_docx_math.repair(source, plan_path, output, None)
            self.assertFalse(output.exists())
            self.assertFalse(list(Path(directory).glob(".wmd-*")))
            with patch.object(repair_docx_math.os, "link", side_effect=OSError("write failure")):
                with self.assertRaises(OSError):
                    repair_docx_math.repair(source, plan_path, output, None)
            self.assertFalse(output.exists())
            self.assertFalse(list(Path(directory).glob(".wmd-*")))

    def test_cross_paragraph_bookmark_is_rejected(self):
        selected = plan()
        selected["operations"] = selected["operations"][1:]
        def change(xml):
            return xml.replace('<w:p><w:r><w:t>勾股', '<w:p><w:bookmarkStart w:id="1" w:name="range"/><w:r><w:t>勾股').replace('<w:sectPr>', '<w:p><w:bookmarkEnd w:id="1"/></w:p><w:sectPr>')
        with self.assertRaises(repair_docx_math.PlanError):
            self.run_fixture(change, selected)

    def test_explicit_out_of_bounds_tab_is_rejected(self):
        selected = plan()
        selected["operations"][1]["right_tab_twips"] = 9360
        with self.assertRaises(repair_docx_math.PlanError):
            self.run_fixture(selected=selected)

    def test_display_does_not_drop_prose_outside_declared_formula(self):
        selected = plan()
        selected["operations"][1]["expected_text"] = "解释文字：" + selected["operations"][1]["expected_text"]
        with self.assertRaises(repair_docx_math.PlanError):
            self.run_fixture(lambda x: x.replace('<w:t>∫', '<w:t>解释文字：∫'), selected)
        selected = plan()
        del selected["operations"][1]["source_formula"]
        with self.assertRaises(repair_docx_math.PlanError):
            self.run_fixture(selected=selected)

    def test_empty_formula_structure_is_rejected(self):
        for empty in ([], "", {"kind": "text", "value": ""}):
            selected = plan()
            selected["operations"][0]["formula"] = empty
            with self.subTest(empty=empty), self.assertRaises(repair_docx_math.PlanError):
                self.run_fixture(selected=selected)

    def test_report_cannot_overwrite_source_or_plan(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            make_docx(source)
            original = source.read_bytes()
            plan_path = Path(directory) / "plan.json"
            plan_path.write_text(json.dumps(plan()), encoding="utf-8")
            for report in (source, plan_path):
                with self.assertRaises(repair_docx_math.PlanError):
                    repair_docx_math.repair(source, plan_path, Path(directory) / "out.docx", report)
            self.assertEqual(source.read_bytes(), original)

    def test_xml_empty_paragraph_span_does_not_consume_following_content(self):
        xml = f'<w:document xmlns:w="{W}"><w:body><w:p w:rsidR="&gt;"/><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'.encode()
        spans = repair_docx_math.paragraph_spans(xml)
        self.assertEqual(xml[slice(*spans[0])], b'<w:p w:rsidR="&gt;"/>')
        self.assertTrue(xml[slice(*spans[1])].endswith(b'</w:p>'))

    def test_bounded_repair_creates_omml_and_tab_aligned_number(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            temporary_path = Path(temporary)
            source = temporary_path / "source.docx"
            output = temporary_path / "converted.docx"
            report_path = temporary_path / "report.json"
            plan_path = temporary_path / "plan.json"
            make_docx(source)
            before = source.read_bytes()
            plan_path.write_text(json.dumps(plan(), ensure_ascii=False), encoding="utf-8")

            report = repair_docx_math.repair(source, plan_path, output, report_path)

            self.assertEqual(source.read_bytes(), before, "原稿不得被脚本修改")
            self.assertTrue(output.exists())
            self.assertEqual(report["omml_before"], 0)
            self.assertEqual(report["omml_after"], 2)
            with zipfile.ZipFile(source) as original, zipfile.ZipFile(output) as converted:
                self.assertEqual(
                    original.read("customXml/item1.xml"), converted.read("customXml/item1.xml")
                )
                document_xml = converted.read("word/document.xml")
                root = ET.fromstring(document_xml)
            root_start = re.search(rb"<w:document\b[^>]*>", document_xml).group(0)
            self.assertIn(b'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', root_start)
            self.assertIn(b'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"', root_start)
            self.assertIn(b'mc:Ignorable="w14 wp14"', root_start)
            paragraphs = list(root.iter(f"{{{W}}}p"))
            self.assertEqual(len(root.findall(".//m:oMath", NS)), 2)
            self.assertEqual(len(paragraphs[1].findall("./w:r/w:tab", NS)), 2)
            tabs = paragraphs[1].findall("./w:pPr/w:tabs/w:tab", NS)
            self.assertEqual([(tab.get(f"{{{W}}}val"), tab.get(f"{{{W}}}pos")) for tab in tabs], [("center", "4320"), ("right", "8640")])
            visible_text = "".join(node.text or "" for node in paragraphs[1].iter(f"{{{W}}}t"))
            self.assertEqual(visible_text, "(1)")
            self.assertNotIn(" " * 10, visible_text)
            self.assertIn("化学式 H₂O 和 CO₂ 保持为可编辑文本。", "".join(node.text or "" for node in paragraphs[2].iter(f"{{{W}}}t")))
            self.assertEqual(json.loads(report_path.read_text(encoding="utf-8"))["operations_applied"][1]["action"], "display")

    def test_refuses_a_stale_or_mismatched_plan(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            temporary_path = Path(temporary)
            source = temporary_path / "source.docx"
            output = temporary_path / "converted.docx"
            plan_path = temporary_path / "plan.json"
            make_docx(source)
            stale = plan()
            stale["operations"][0]["expected_text"] = "已被别人改过"
            plan_path.write_text(json.dumps(stale, ensure_ascii=False), encoding="utf-8")
            with self.assertRaises(repair_docx_math.PlanError):
                repair_docx_math.repair(source, plan_path, output, None)
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
