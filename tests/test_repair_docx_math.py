from __future__ import annotations

import importlib.util
import json
import re
import tempfile
import unittest
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
  <w:sectPr/>
</w:body></w:document>'''.encode("utf-8")
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("word/document.xml", document)
        archive.writestr("customXml/item1.xml", b"unchanged-by-repair")


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
            self.assertEqual([(tab.get(f"{{{W}}}val"), tab.get(f"{{{W}}}pos")) for tab in tabs], [("center", "4680"), ("right", "9360")])
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
