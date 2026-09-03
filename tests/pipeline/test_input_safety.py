import sys
import tempfile
import unittest
import warnings
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from input_safety import validate_pptx_archive, validate_print_play_counts  # noqa: E402


class InputSafetyTests(unittest.TestCase):
    def make_archive(self, entries):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        path = Path(temp_dir.name) / "playbook.pptx"
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, body in entries.items():
                archive.writestr(name, body)
        return path

    def test_accepts_minimal_pptx_container(self):
        path = self.make_archive({
            "[Content_Types].xml": "<Types/>",
            "ppt/presentation.xml": "<presentation/>",
        })
        validate_pptx_archive(path)

    def test_rejects_missing_powerpoint_members(self):
        path = self.make_archive({"notes.txt": "not a presentation"})
        with self.assertRaisesRegex(ValueError, "valid PowerPoint"):
            validate_pptx_archive(path)

    def test_rejects_unsafe_member_paths(self):
        path = self.make_archive({
            "[Content_Types].xml": "<Types/>",
            "ppt/presentation.xml": "<presentation/>",
            "../escape.bin": "x",
        })
        with self.assertRaisesRegex(ValueError, "unsafe path"):
            validate_pptx_archive(path)

    def test_rejects_duplicate_archive_entries(self):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        path = Path(temp_dir.name) / "playbook.pptx"
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("[Content_Types].xml", "<Types/>")
                archive.writestr("ppt/presentation.xml", "<presentation/>")
                archive.writestr("ppt/presentation.xml", "<duplicate/>")
        with self.assertRaisesRegex(ValueError, "duplicate entries"):
            validate_pptx_archive(path)

    def test_rejects_print_counts_above_capacity(self):
        plays = [{"section": "OFFENSE"} for _ in range(65)]
        with self.assertRaisesRegex(ValueError, "at most 64"):
            validate_print_play_counts(plays)

        plays = [{"section": "DEFENSE"} for _ in range(25)]
        with self.assertRaisesRegex(ValueError, "at most 24"):
            validate_print_play_counts(plays)

    def test_accepts_capacity_boundaries(self):
        plays = (
            [{"section": "OFFENSE"} for _ in range(64)]
            + [{"section": "DEFENSE"} for _ in range(24)]
        )
        validate_print_play_counts(plays)


if __name__ == "__main__":
    unittest.main()
