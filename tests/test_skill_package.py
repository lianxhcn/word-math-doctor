from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILL_FILE = ROOT / "skills" / "word-math-doctor" / "SKILL.md"
EXPECTED_VERSION = "0.4.2-beta"


class SkillPackageTests(unittest.TestCase):
    def read_front_matter(self) -> str:
        content = SKILL_FILE.read_text(encoding="utf-8")
        match = re.match(r"^---\n(.*?)\n---\n", content, re.DOTALL)
        self.assertIsNotNone(match, "SKILL.md 必须以 YAML front matter 开始")
        return match.group(1) if match else ""

    def test_front_matter_has_clear_trigger_contract(self) -> None:
        front_matter = self.read_front_matter()
        self.assertRegex(front_matter, r"(?m)^name: word-math-doctor$")
        description = re.search(r"(?m)^description: (.+)$", front_matter)
        self.assertIsNotNone(description)
        self.assertLessEqual(len(description.group(1)), 1024)

    def test_portable_manifests_match_release(self) -> None:
        manifest_paths = [
            ROOT / "plugin.json",
            ROOT / ".codex-plugin" / "plugin.json",
            ROOT / ".claude-plugin" / "plugin.json",
        ]
        for path in manifest_paths:
            manifest = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["name"], "word-math-doctor")
            self.assertEqual(manifest["version"], EXPECTED_VERSION)
            self.assertEqual(manifest["license"], "CC-BY-NC-4.0")

        codex_marketplace = json.loads(
            (ROOT / ".agents" / "plugins" / "marketplace.json").read_text(encoding="utf-8")
        )
        self.assertEqual(codex_marketplace["plugins"][0]["name"], "word-math-doctor")

        claude_marketplace = json.loads(
            (ROOT / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8")
        )
        self.assertEqual(claude_marketplace["plugins"][0]["name"], "word-math-doctor")
        self.assertEqual(claude_marketplace["plugins"][0]["version"], EXPECTED_VERSION)

    def test_readme_documents_supported_install_paths(self) -> None:
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        for phrase in [
            "npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent codex",
            "npx skills add lianxhcn/word-math-doctor --skill word-math-doctor --agent claude-code",
            "codex plugin marketplace add",
            "claude plugin marketplace add",
            "CC BY-NC 4.0",
        ]:
            self.assertIn(phrase, readme)


if __name__ == "__main__":
    unittest.main()
