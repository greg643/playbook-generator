import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from playbook_pipeline import PlaybookGenerator  # noqa: E402


class GeneratorContractTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        root = Path(self.temp_dir.name)
        self.images = root / "images"
        self.output = root / "output"
        self.images.mkdir()
        Image.new("RGB", (160, 120), "white").save(self.images / "01.png")

    def generator(self):
        return PlaybookGenerator(self.images, self.output)

    def test_produces_exact_requested_output_set(self):
        produced = self.generator().generate_all(
            gen_offense=True,
            gen_defense=False,
            offense_coach_card=True,
            offense_wristband=False,
            defense_coach_card=False,
            defense_wristband=False,
        )
        self.assertEqual(produced, ["offense_coach_card.pdf"])
        self.assertEqual(
            {path.name for path in self.output.glob("*.pdf")},
            {"offense_coach_card.pdf"},
        )

    def test_single_section_deck_produces_available_outputs(self):
        # All four outputs requested (the upload page default) against an
        # offense-only deck: the offense pair is produced, defense skipped.
        produced = self.generator().generate_all(
            gen_offense=True,
            gen_defense=True,
            offense_coach_card=True,
            offense_wristband=True,
            defense_coach_card=True,
            defense_wristband=True,
        )
        self.assertEqual(produced, ["offense_coach_card.pdf", "offense_wristband.pdf"])

    def test_nothing_producible_is_an_error(self):
        with self.assertRaisesRegex(ValueError, "No plays were found"):
            self.generator().generate_all(
                gen_offense=False,
                gen_defense=True,
                offense_coach_card=False,
                offense_wristband=False,
                defense_coach_card=True,
                defense_wristband=False,
            )

    def test_unrelated_pdf_in_output_dir_is_ignored(self):
        self.output.mkdir()
        (self.output / "notes.pdf").write_bytes(b"%PDF-1.4 unrelated")
        produced = self.generator().generate_all(
            gen_offense=True,
            gen_defense=False,
            offense_coach_card=True,
            offense_wristband=False,
            defense_coach_card=False,
            defense_wristband=False,
        )
        self.assertEqual(produced, ["offense_coach_card.pdf"])
        self.assertTrue((self.output / "notes.pdf").exists())

    def test_rejects_images_outside_print_capacity(self):
        Image.new("RGB", (32, 32), "white").save(self.images / "17.png")
        with self.assertRaisesRegex(ValueError, "Unsupported play image filename"):
            self.generator().load_images()

    def test_rejects_duplicate_numeric_aliases(self):
        Image.new("RGB", (32, 32), "white").save(self.images / "1.jpg")
        with self.assertRaisesRegex(ValueError, "Duplicate play image slot"):
            self.generator().load_images()


if __name__ == "__main__":
    unittest.main()
