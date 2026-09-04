from __future__ import annotations

import json
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL_DIRECTORY = ROOT / "skills" / "word-math-doctor"
SKILL_FILE = SKILL_DIRECTORY / "SKILL.md"
PLUGIN_FILE = ROOT / ".codex-plugin" / "plugin.json"
EXPECTED_VERSION = "0.2.0-beta"
EXPECTED_LICENSE = "CC-BY-NC-4.0"


class SkillPackageTests(unittest.TestCase):
    def read_front_matter(self) -> str:
        content = SKILL_FILE.read_text(encoding="utf-8")
        match = re.match(r"^---\n(.*?)\n---\n", content, re.DOTALL)
        self.assertIsNotNone(match, "SKILL.md 必须以 YAML front matter 开始")
        return match.group(1) if match else ""

    def test_front_matter_has_publishable_identity(self) -> None:
        front_matter = self.read_front_matter()

        self.assertRegex(front_matter, r"(?m)^name: word-math-doctor$")
        self.assertRegex(front_matter, r"(?m)^license: CC-BY-NC-4.0$")
        self.assertRegex(front_matter, r'(?m)^  version: "0\.2\.0-beta"$')
        self.assertRegex(front_matter, r"(?m)^  compatibility: .+$")

        description = re.search(r"(?m)^description: (.+)$", front_matter)
        self.assertIsNotNone(description)
        self.assertLessEqual(len(description.group(1)), 1024)

    def test_local_skill_references_exist(self) -> None:
        content = SKILL_FILE.read_text(encoding="utf-8")
        references = re.findall(r"\[[^\]]+\]\(([^)#]+)(?:#[^)]*)?\)", content)

        for reference in references:
            if "://" in reference or reference.startswith("/"):
                continue
            path = (SKILL_DIRECTORY / reference).resolve()
            self.assertTrue(path.is_file(), f"缺少 SKILL.md 引用的本地文件：{reference}")

    def test_plugin_manifest_matches_skill_release(self) -> None:
        manifest = json.loads(PLUGIN_FILE.read_text(encoding="utf-8"))

        self.assertEqual(manifest["name"], "word-math-doctor")
        self.assertEqual(manifest["version"], EXPECTED_VERSION)
        self.assertEqual(manifest["license"], EXPECTED_LICENSE)
        self.assertEqual(manifest["skills"], "./skills/")

    def test_readme_contains_direct_install_and_license_guidance(self) -> None:
        readme = (ROOT / "README.md").read_text(encoding="utf-8")

        self.assertIn(
            "npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent codex",
            readme,
        )
        self.assertIn("CC BY-NC 4.0", readme)
        self.assertIn("skills.sh", readme)

    def test_public_release_documents_exist(self) -> None:
        self.assertTrue((ROOT / "SECURITY.md").is_file())
        self.assertTrue((ROOT / "CITATION.cff").is_file())


if __name__ == "__main__":
    unittest.main()
