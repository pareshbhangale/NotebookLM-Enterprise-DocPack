from __future__ import annotations

import io
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import notebooklm_enterprise_docpack


class NotebookLMEnterpriseDocPackTests(unittest.TestCase):
    def test_default_paths_are_generic(self) -> None:
        self.assertEqual(notebooklm_enterprise_docpack.DEFAULT_MANIFEST.name, "manifest.csv")
        self.assertEqual(notebooklm_enterprise_docpack.DEFAULT_MASTER_MD.name, "knowledge_pack.md")
        self.assertEqual(notebooklm_enterprise_docpack.DEFAULT_DOCX.name, "docx_packs")

    def test_read_manifest_accepts_url_alias_and_title_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            manifest = Path(tmp_dir) / "sources.csv"
            manifest.write_text(
                "type,name,source_url\n"
                "pdf,Policy guide,https://example.com/policy.pdf\n"
                "zip,,https://example.com/reference.zip\n",
                encoding="utf-8",
            )

            self.assertEqual(
                notebooklm_enterprise_docpack.read_manifest_csv(manifest),
                [
                    ("pdf", "Policy guide", "https://example.com/policy.pdf"),
                    ("zip", "reference.zip", "https://example.com/reference.zip"),
                ],
            )

    def test_missing_default_manifest_explains_how_to_start(self) -> None:
        missing_manifest = PROJECT_ROOT / "Input" / "not-present.csv"

        with self.assertRaisesRegex(FileNotFoundError, "example_download_manifest.csv"):
            notebooklm_enterprise_docpack.read_manifest_csv(missing_manifest)

    def test_preflight_rejects_missing_reused_stage_input(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)

            with self.assertRaisesRegex(FileNotFoundError, "Raw input folder not found"):
                notebooklm_enterprise_docpack.validate_stage_inputs(
                    manifest=root / "manifest.csv",
                    raw_dir=root / "raw",
                    md_dir=root / "markdown",
                    master_md=root / "knowledge_pack.md",
                    skip_download=True,
                    skip_convert=False,
                    skip_merge=False,
                    skip_docx=False,
                )

    def test_zip_extraction_skips_path_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = Path(tmp_dir) / "extracted"
            outside = Path(tmp_dir) / "outside.txt"
            archive_bytes = io.BytesIO()
            with zipfile.ZipFile(archive_bytes, "w") as archive:
                archive.writestr("guides/overview.txt", "safe content")
                archive.writestr("../outside.txt", "unsafe content")

            archive_bytes.seek(0)
            with zipfile.ZipFile(archive_bytes) as archive:
                notebooklm_enterprise_docpack.extract_zip_safely(archive, target)

            self.assertEqual((target / "guides" / "overview.txt").read_text(), "safe content")
            self.assertFalse(outside.exists())

    def test_source_blocks_round_trip_and_split_at_word_limit(self) -> None:
        master = (
            "---\n# FILE: guides/one.md\n---\n\nfirst second third\n\n"
            "---\n# FILE: guides/two.md\n---\n\nfour fifth"
        )

        self.assertEqual(
            notebooklm_enterprise_docpack.parse_master_md(master),
            [("guides/one.md", "first second third"), ("guides/two.md", "four fifth")],
        )
        self.assertEqual(
            notebooklm_enterprise_docpack.split_text_by_words("one two three four", 2),
            ["one two", "three four"],
        )

    @unittest.skipIf(notebooklm_enterprise_docpack.Document is None, "python-docx is not installed")
    def test_docx_pack_uses_generic_filename(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_dir = Path(tmp_dir) / "packs"
            pack_paths = notebooklm_enterprise_docpack.build_docx_pack(
                blocks=[("guides/overview.md", "# Overview\n\nUseful content")],
                out_dir=out_dir,
                master_md=Path(tmp_dir) / "knowledge_pack.md",
                chunk_words=100,
                pack_words=100,
                max_words=200,
                max_docx_parts=0,
            )

            self.assertEqual([path.name for path in pack_paths], ["knowledge_pack_part_001.docx"])


if __name__ == "__main__":
    unittest.main()
