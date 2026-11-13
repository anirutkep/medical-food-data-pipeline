import json
import re
import tomllib
import unittest
from pathlib import Path


class TestPackagingAndManifests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pkg_root = Path(__file__).resolve().parent.parent

    def test_root_files_exist(self):
        """Root files must exist: LICENSE, .gitignore, requirements.txt, pyproject.toml, README.md."""
        required_files = [
            "LICENSE",
            ".gitignore",
            "requirements.txt",
            "pyproject.toml",
            "README.md",
        ]
        for filename in required_files:
            file_path = self.pkg_root / filename
            self.assertTrue(
                file_path.is_file(),
                f"Missing required root file: {filename} at {file_path}",
            )

    def test_json_manifests(self):
        """Validate all 4 agent/plugin JSON manifests."""
        manifest_specs = [
            (
                self.pkg_root / ".claude-plugin" / "plugin.json",
                True,   # is_plugin
                False,  # is_marketplace
            ),
            (
                self.pkg_root / ".claude-plugin" / "marketplace.json",
                False,
                True,
            ),
            (
                self.pkg_root / ".codex-plugin" / "plugin.json",
                True,
                False,
            ),
            (
                self.pkg_root / ".agents" / "plugins" / "marketplace.json",
                False,
                True,
            ),
        ]

        for path, is_plugin, is_marketplace in manifest_specs:
            self.assertTrue(
                path.is_file(),
                f"Manifest file missing: {path.relative_to(self.pkg_root)}",
            )
            raw_text = path.read_text(encoding="utf-8")
            data = json.loads(raw_text)

            # Verify real author/repo metadata and no leftover placeholders
            self.assertNotIn("<AUTHOR_NAME>", raw_text)
            self.assertNotIn("your-org/medical-food-data-pipeline", raw_text)
            self.assertTrue(
                "Anirut Kepan, RD" in raw_text or "anirutkep/medical-food-data-pipeline" in raw_text,
                f"Manifest {path.name} must contain author or repo url",
            )

            if is_plugin:
                self.assertEqual(data.get("name"), "medical-food-data-pipeline")
                self.assertEqual(data.get("skills"), "./skills/")
                self.assertEqual(data.get("license"), "MIT")
                self.assertEqual(data.get("version"), "0.1.0")

            if is_marketplace:
                self.assertEqual(data.get("name"), "medical-food-data-pipeline-marketplace")
                self.assertIn("owner", data)
                self.assertIn("plugins", data)
                self.assertIsInstance(data["plugins"], list)
                self.assertGreaterEqual(len(data["plugins"]), 1)
                entry = data["plugins"][0]
                self.assertEqual(entry.get("name"), "medical-food-data-pipeline")
                self.assertEqual(entry.get("source"), "./")
                self.assertEqual(entry.get("skills"), "./skills/")
                self.assertEqual(entry.get("license"), "MIT")

    def test_pyproject_toml(self):
        """Verify pyproject.toml conforms to PEP 621 specifications."""
        pyproject_path = self.pkg_root / "pyproject.toml"
        self.assertTrue(pyproject_path.is_file(), "pyproject.toml missing")

        raw_text = pyproject_path.read_text(encoding="utf-8")
        data = tomllib.loads(raw_text)

        project = data.get("project", {})
        self.assertEqual(project.get("name"), "medical-food-data-pipeline-skill")
        self.assertEqual(project.get("version"), "0.1.0")
        self.assertEqual(project.get("requires-python"), ">=3.10")
        self.assertIn("openpyxl>=3.1.0", project.get("dependencies", []))
        self.assertNotIn("<AUTHOR_NAME>", raw_text)
        self.assertNotIn("your-org/medical-food-data-pipeline", raw_text)
        self.assertIn("Anirut Kepan, RD", raw_text)
        self.assertIn("anirutkep/medical-food-data-pipeline", raw_text)

        # Optional dependencies check
        opt_deps = project.get("optional-dependencies", {})
        self.assertIn("ocr", opt_deps)
        self.assertIn("dev", opt_deps)

    def test_requirements_txt(self):
        """Verify requirements.txt contains openpyxl>=3.1.0."""
        req_path = self.pkg_root / "requirements.txt"
        self.assertTrue(req_path.is_file(), "requirements.txt missing")
        content = req_path.read_text(encoding="utf-8")
        self.assertIn("openpyxl>=3.1.0", content)

    def test_license(self):
        """Verify MIT License contains required copyright notice."""
        lic_path = self.pkg_root / "LICENSE"
        self.assertTrue(lic_path.is_file(), "LICENSE missing")
        content = lic_path.read_text(encoding="utf-8")
        self.assertIn("MIT License", content)
        self.assertNotIn("<AUTHOR_NAME>", content)
        self.assertIn("Anirut Kepan, RD", content)
        self.assertIn("2026", content)

    def test_gitignore(self):
        """Verify .gitignore contains expected patterns."""
        gi_path = self.pkg_root / ".gitignore"
        self.assertTrue(gi_path.is_file(), ".gitignore missing")
        content = gi_path.read_text(encoding="utf-8")
        required_patterns = [
            "__pycache__/",
            "*.py[cod]",
            ".DS_Store",
            "Thumbs.db",
            ".venv/",
            ".ocr_cache/",
            "dist/",
            "*.egg-info/",
        ]
        for pattern in required_patterns:
            self.assertIn(pattern, content)

    def test_readme_bilingual_and_keywords(self):
        """Verify README.md covers required keywords, install guides, and bilingual text."""
        readme_path = self.pkg_root / "README.md"
        self.assertTrue(readme_path.is_file(), "README.md missing")
        content = readme_path.read_text(encoding="utf-8")

        required_keywords = [
            "MOPH",
            "445",
            "Windows",
            "macOS",
            "npx skills add",
            "Codex",
            "Claude Code",
            "TH Sarabun New",
            "Cordia New",
            "openpyxl",
        ]
        for kw in required_keywords:
            self.assertIn(kw, content, f"README.md missing required keyword: {kw}")

        # Bilingual check: ensure Thai characters exist alongside English
        has_thai = any("\u0e00" <= ch <= "\u0e7f" for ch in content)
        self.assertTrue(has_thai, "README.md must contain Thai text for bilingual documentation")


class TestSkillMdAndStylingPortability(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pkg_root = Path(__file__).resolve().parent.parent
        cls.skill_dir = cls.pkg_root / "skills" / "medical-food-data-pipeline"
        cls.skill_md = cls.skill_dir / "SKILL.md"
        cls.styling_guide = cls.skill_dir / "references" / "excel_styling_guide.md"

    def test_skill_md_frontmatter_and_keywords(self):
        """SKILL.md exists, has valid YAML frontmatter, and contains required keywords."""
        self.assertTrue(self.skill_md.is_file(), f"SKILL.md not found at {self.skill_md}")
        content = self.skill_md.read_text(encoding="utf-8")

        # Must have YAML frontmatter delimiters
        self.assertTrue(content.startswith("---"), "SKILL.md must start with YAML frontmatter '---'")
        parts = content.split("---", 2)
        self.assertGreaterEqual(len(parts), 3, "SKILL.md must have valid closing '---' for frontmatter")
        frontmatter = parts[1]
        self.assertIn("name: medical-food-data-pipeline", frontmatter)
        self.assertIn("description:", frontmatter)

        # Check required keywords (MOPH, 1000 kcal, enteral nutrition, อาหารทางการแพทย์)
        content_lower = content.lower()
        self.assertIn("moph", content_lower)
        self.assertIn("1000 kcal", content_lower)
        self.assertIn("enteral nutrition", content_lower)
        self.assertIn("อาหารทางการแพทย์", content)

    def test_skill_md_no_harness_leakage(self):
        """SKILL.md must not leak harness tools, developer paths, or POSIX-only destructive commands."""
        self.assertTrue(self.skill_md.is_file(), f"SKILL.md not found at {self.skill_md}")
        content = self.skill_md.read_text(encoding="utf-8")
        forbidden = [
            "swarm{",
            "read_file",
            "/Users/anirutkep",
            "-exec rm -rf",
            '-name ".DS_Store" -delete',
        ]
        for term in forbidden:
            self.assertNotIn(term, content, f"SKILL.md contains forbidden leaked substring: '{term}'")

    def test_skill_md_references_scripts_and_assets(self):
        """SKILL.md explicitly references relative-to-<skill_dir> scripts, templates, and guides."""
        self.assertTrue(self.skill_md.is_file(), f"SKILL.md not found at {self.skill_md}")
        content = self.skill_md.read_text(encoding="utf-8")
        required_refs = [
            "<skill_dir>/scripts/ocr_row_extractor.py",
            "<skill_dir>/scripts/unit_normalizer.py",
            "<skill_dir>/scripts/apply_unit_normalization.py",
            "<skill_dir>/scripts/pipeline_audit.py",
            "<skill_dir>/scripts/cleanup_workspace.py",
            "<skill_dir>/assets/brand_data_template.py",
            "<skill_dir>/references/excel_styling_guide.md",
        ]
        for ref in required_refs:
            self.assertIn(ref, content, f"SKILL.md missing reference: '{ref}'")

    def test_excel_styling_guide_portability(self):
        """references/excel_styling_guide.md exists and covers fonts, padding margin, and cross-platform renderers."""
        self.assertTrue(self.styling_guide.is_file(), f"Styling guide not found at {self.styling_guide}")
        content = self.styling_guide.read_text(encoding="utf-8")
        required_terms = [
            "TH Sarabun New",
            "Cordia New",
            "+5",
            "macOS",
            "LibreOffice",
        ]
        for term in required_terms:
            self.assertIn(term, content, f"excel_styling_guide.md missing required term: '{term}'")

    def test_no_commercial_brands_media_or_pii_leaked(self):
        """Stage 1: Verify zero media files, commercial product names, or developer PII exist in published artifacts."""
        media_exts = {".jpg", ".jpeg", ".png", ".pdf", ".webp", ".heic"}
        text_exts = {".md", ".py", ".toml", ".json", ".txt", ""}
        forbidden_patterns = [
            r"\bAbbott\b",
            r"\bNestle\b",
            r"\bOtsuka\b",
            r"\bDanone\b",
            r"\bJusmin\b",
            r"\bGlucerna\b",
            r"\bNepro\b",
            r"\bProSure\b",
            r"\bPeptamen\b",
            r"\bIsocal\b",
            r"\bBlendera\b",
            r"\bNeo-Mune\b",
            r"\bOral Impact\b",
            r"\bInfatrini\b",
            r"\bMabeedee\b",
            r"\bJevity\b",
            r"\bPediasure\b",
            r"\bPulmocare\b",
            r"\bAlitraq\b",
            r"\bOptifibre\b",
            r"\bPan-Enteral\b",
            r"\bGen-DM\b",
            r"/Users/",
            r"M2570",
            r"<AUTHOR_NAME>",
            r"your-org/",
        ]
        compiled = [re.compile(p, re.IGNORECASE) for p in forbidden_patterns]

        for p in self.pkg_root.rglob("*"):
            if not p.is_file():
                continue
            rel = p.relative_to(self.pkg_root)
            if "__pycache__" in rel.parts or rel.parts[0] == "tests":
                continue
            self.assertNotIn(
                p.suffix.lower(),
                media_exts,
                f"Published package must not contain media/PDF evidence: {rel}",
            )
            if p.suffix.lower() in text_exts:
                txt = p.read_text(encoding="utf-8")
                for rx in compiled:
                    self.assertIsNone(
                        rx.search(txt),
                        f"Published file '{rel}' matched forbidden pattern '{rx.pattern}'",
                    )

    def test_readme_relative_links_and_metadata_consistency(self):
        """Stage 2: Verify README relative links exist and author/version/repo match across all release files."""
        readme_text = (self.pkg_root / "README.md").read_text(encoding="utf-8")
        # Strip fenced code blocks and check markdown links [label](target)
        stripped = re.sub(r"```.*?```", "", readme_text, flags=re.DOTALL)
        for match in re.finditer(r"\[[^\]]*\]\(([^)]+)\)", stripped):
            target = match.group(1).strip()
            if target.startswith(("http://", "https://", "#", "mailto:")):
                continue
            resolved = (self.pkg_root / target.split("#", 1)[0]).resolve()
            self.assertTrue(
                resolved.exists(),
                f"README.md broken relative link: '{target}' -> {resolved}",
            )

        self.assertIn("Anirut Kepan, RD", readme_text)
        self.assertIn("anirutkep/medical-food-data-pipeline", readme_text)
        self.assertIn("0.1.0", readme_text)


if __name__ == "__main__":
    unittest.main()
