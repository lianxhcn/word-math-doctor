from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCANNER = ROOT / "skills" / "word-math-doctor" / "scripts" / "scan_docx_math.py"

DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p><w:r><w:t>Normal text with \\alpha and x=1 and b_i²x₁</w:t></w:r></w:p>
    <m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>
    <w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:commentRangeStart w:id="0"/>
    <w:sectPr/>
  </w:body>
</w:document>"""

COMMENTS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>Fix b_i</w:t></w:r></w:p></w:comment>
</w:comments>"""


class ScanDocxMathTests(unittest.TestCase):
    def build_fixture(self, directory: Path) -> Path:
        document = directory / "fixture.docx"
        with zipfile.ZipFile(document, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("word/document.xml", DOCUMENT_XML)
            archive.writestr("word/comments.xml", COMMENTS_XML)
            archive.writestr("word/media/image1.png", b"not-a-real-png")
            archive.writestr("word/embeddings/Equation.1.bin", b"ole")
        return document

    def run_scanner(
        self,
        document: Path,
        output: Path,
        baseline: Path | None = None,
        extra_args: list[str] | None = None,
    ) -> dict:
        command = [sys.executable, str(SCANNER), str(document), "--output", str(output)]
        if baseline:
            command.extend(["--baseline", str(baseline)])
        if extra_args:
            command.extend(extra_args)
        subprocess.run(command, check=True, capture_output=True, text=True)
        return json.loads(output.read_text(encoding="utf-8"))

    def test_inventory_and_baseline(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            document = self.build_fixture(directory)
            first = self.run_scanner(document, directory / "first.json")
            counts = first["counts"]

            self.assertEqual(first["schema"], "word-math-doctor/docx-math-audit/v2")
            self.assertEqual(first["privacy"]["source_paths"], "omitted")
            self.assertEqual(first["privacy"]["candidate_text_snippets"], "omitted")
            self.assertNotIn("path", first["source"])
            self.assertNotIn(
                "text_math_candidate_snippets", first["content_parts"][0]
            )
            self.assertEqual(counts["omml_objects"], 1)
            self.assertEqual(counts["omml_paragraphs"], 1)
            self.assertEqual(counts["tables"], 1)
            self.assertEqual(counts["table_cells"], 1)
            self.assertEqual(counts["legacy_superscript_runs"], 1)
            self.assertGreaterEqual(counts["unicode_superscript_or_subscript_characters"], 2)
            self.assertGreaterEqual(counts["visible_latex_candidates"], 1)
            self.assertEqual(counts["comments"], 1)
            self.assertEqual(counts["media_files"], 1)
            self.assertEqual(counts["embedded_object_files"], 1)

            second = self.run_scanner(document, directory / "second.json", directory / "first.json")
            self.assertEqual(second["baseline_comparison"]["count_differences"], {})
            self.assertNotIn("baseline_path", second["baseline_comparison"])

            detailed = self.run_scanner(
                document,
                directory / "detailed.json",
                extra_args=["--include-snippets", "--include-paths"],
            )
            self.assertEqual(detailed["source"]["path"], str(document))
            document_part = next(
                part
                for part in detailed["content_parts"]
                if part["part"] == "word/document.xml"
            )
            self.assertTrue(
                document_part["latex_candidate_snippets"]
            )


if __name__ == "__main__":
    unittest.main()
