import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from playbook_pipeline import crop_plays  # noqa: E402


class CropPlayTests(unittest.TestCase):
    def test_missing_rendered_slides_is_terminal(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(RuntimeError, "No slide images"):
                crop_plays(
                    [
                        {
                            "slide_index": 0,
                            "filename": "01.png",
                            "play_id": "1",
                            "play_name": "Test",
                            "crop_box_emu": (0, 0, 1, 1),
                        }
                    ],
                    [],
                    1,
                    1,
                    Path(directory),
                )

    def test_large_render_crop_is_bounded_before_generator_loading(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            slide_path = root / "slide-01.png"
            output_dir = root / "plays"
            Image.new("RGB", (2400, 1200), "white").save(slide_path)

            saved = crop_plays(
                [
                    {
                        "slide_index": 0,
                        "filename": "01.png",
                        "play_id": "1",
                        "play_name": "Test",
                        "crop_box_emu": (0, 0, 2400, 1200),
                    }
                ],
                [slide_path],
                2400,
                1200,
                output_dir,
            )

            self.assertEqual(saved, [output_dir / "01.png"])
            with Image.open(saved[0]) as image:
                self.assertEqual(image.size, (1800, 900))

    def test_duplicate_output_names_are_terminal_before_writing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            slide_path = root / "slide-01.png"
            output_dir = root / "plays"
            Image.new("RGB", (100, 100), "white").save(slide_path)
            duplicate = {
                "slide_index": 0,
                "filename": "01.png",
                "play_id": "1",
                "play_name": "Test",
                "crop_box_emu": (0, 0, 100, 100),
            }

            with self.assertRaisesRegex(RuntimeError, "duplicate output filenames"):
                crop_plays(
                    [duplicate, dict(duplicate)],
                    [slide_path],
                    100,
                    100,
                    output_dir,
                )

            self.assertEqual(list(output_dir.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
