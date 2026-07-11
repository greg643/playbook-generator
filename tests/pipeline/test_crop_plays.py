import sys
import tempfile
import unittest
from pathlib import Path


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


if __name__ == "__main__":
    unittest.main()
