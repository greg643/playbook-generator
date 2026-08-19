import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from playbook_pipeline import (  # noqa: E402
    PlaybookGenerator,
    wristband_positions,
    wristband_title_allowed,
)


def page_text(pdf_path, page=0):
    from pypdf import PdfReader

    return PdfReader(str(pdf_path)).pages[page].extract_text()


def image_xobject_count(page):
    """Count distinct embedded images referenced by one PDF page."""
    resources = page["/Resources"].get_object()
    xobjects = resources.get("/XObject", {})
    return sum(
        1
        for reference in xobjects.values()
        if reference.get_object().get("/Subtype") == "/Image"
    )


def xobject_draw_count(page):
    """Count image/form paint operations in one ReportLab page stream."""
    return page.get_contents().get_data().split().count(b"Do")


class WristbandLayoutTests(unittest.TestCase):
    def test_count_adaptive_shapes(self):
        # (count) -> (cards on top row, vertically centered cards, bottom row)
        expected = {
            1: (0, 1, 0),
            2: (0, 2, 0),
            3: (0, 3, 0),
            4: (2, 0, 2),
            5: (2, 1, 2),   # 2-1-2 dice
            6: (3, 0, 3),   # 3 over 3
            7: (4, 0, 3),   # 4 over 3
            8: (4, 0, 4),   # classic 4x4 over two rows
        }
        for n, (top, mid, bottom) in expected.items():
            positions = wristband_positions(n)
            self.assertEqual(len(positions), n)
            rows = [row for _col, row in positions]
            self.assertEqual(
                (rows.count(0), rows.count(0.5), rows.count(1)),
                (top, mid, bottom),
                f"layout shape for {n} plays",
            )

    def test_seven_bottom_row_is_centered(self):
        positions = wristband_positions(7)
        bottom = sorted(col for col, row in positions if row == 1)
        self.assertEqual(bottom, [0.5, 1.5, 2.5])

    def test_column_major_preserves_defense_reading_order(self):
        self.assertEqual(
            wristband_positions(4, column_major=True),
            [(0, 0), (0, 1), (1, 0), (1, 1)],
        )

    def test_title_allowed_only_below_seven(self):
        self.assertTrue(wristband_title_allowed(6))
        self.assertFalse(wristband_title_allowed(7))
        self.assertFalse(wristband_title_allowed(8))


class WristbandTitleTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        root = Path(self.temp_dir.name)
        self.images = root / "images"
        self.output = root / "output"
        self.images.mkdir()

    def make_images(self, offense=0, defense=0):
        for i in range(1, offense + 1):
            Image.new("RGB", (160, 120), "white").save(self.images / f"{i:02d}.png")
        for i in range(1, defense + 1):
            Image.new("RGB", (160, 120), "gray").save(self.images / f"D{i}.png")
        return PlaybookGenerator(self.images, self.output)

    def test_defense_title_shown_by_default(self):
        gen = self.make_images(defense=4)
        gen.create_wristband_sheet_defense(gen.load_images()[1])
        self.assertIn("DEFENSE", page_text(self.output / "defense_wristband.pdf"))

    def test_defense_title_can_be_disabled(self):
        gen = self.make_images(defense=4)
        gen.create_wristband_sheet_defense(gen.load_images()[1], show_title=False)
        self.assertNotIn("DEFENSE", page_text(self.output / "defense_wristband.pdf"))

    def test_offense_title_opt_in_when_it_fits(self):
        gen = self.make_images(offense=5)
        gen.create_wristband_sheet_offense(gen.load_images()[0], show_title=True)
        self.assertIn("OFFENSE", page_text(self.output / "offense_wristband.pdf"))

    def test_offense_title_suppressed_on_full_groups(self):
        # 8-card groups fill the cut-out width: no room for the title even
        # when requested.
        gen = self.make_images(offense=8)
        gen.create_wristband_sheet_offense(gen.load_images()[0], show_title=True)
        self.assertNotIn("OFFENSE", page_text(self.output / "offense_wristband.pdf"))


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

    def test_offense_outputs_paginate_without_dropping_plays(self):
        from pypdf import PdfReader

        # Give every play distinct image bytes so PDF resource counts prove
        # that every logical play, including #17, reached its expected page.
        for index in range(1, 18):
            color = (index, (index * 7) % 256, (index * 13) % 256)
            Image.new("RGB", (32, 24), color).save(self.images / f"{index:02d}.png")

        produced = self.generator().generate_all(
            gen_offense=True,
            gen_defense=False,
            offense_coach_card=True,
            offense_wristband=True,
            defense_coach_card=False,
            defense_wristband=False,
        )

        self.assertEqual(produced, ["offense_coach_card.pdf", "offense_wristband.pdf"])
        coach_pages = PdfReader(str(self.output / "offense_coach_card.pdf")).pages
        wristband_pages = PdfReader(str(self.output / "offense_wristband.pdf")).pages
        self.assertEqual(len(coach_pages), 2)
        self.assertEqual(len(wristband_pages), 3)
        self.assertEqual([image_xobject_count(page) for page in coach_pages], [16, 1])
        self.assertEqual([xobject_draw_count(page) for page in coach_pages], [16, 1])
        self.assertEqual([image_xobject_count(page) for page in wristband_pages], [8, 8, 1])
        self.assertEqual([xobject_draw_count(page) for page in wristband_pages], [48, 48, 6])
        self.assertTrue(all("OFFENSE" in page.extract_text() for page in coach_pages))

    def test_loads_full_sixty_four_play_capacity(self):
        from pypdf import PdfReader

        for index in range(1, 65):
            Image.new("RGB", (4, 4), (index, 0, 0)).save(
                self.images / f"{index:02d}.png"
            )

        generator = self.generator()
        offense, defense = generator.load_images()

        self.assertEqual(len(offense), 64)
        self.assertEqual(defense, [])

        generator.create_coach_card_offense(offense)
        generator.create_wristband_sheet_offense(offense)
        coach_pages = PdfReader(str(self.output / "offense_coach_card.pdf")).pages
        wristband_pages = PdfReader(str(self.output / "offense_wristband.pdf")).pages
        self.assertEqual(len(coach_pages), 4)
        self.assertEqual(len(wristband_pages), 8)
        self.assertEqual([image_xobject_count(page) for page in coach_pages], [16] * 4)
        self.assertEqual([xobject_draw_count(page) for page in coach_pages], [16] * 4)
        self.assertEqual([image_xobject_count(page) for page in wristband_pages], [8] * 8)
        self.assertEqual([xobject_draw_count(page) for page in wristband_pages], [48] * 8)

    def test_aggregate_budget_uses_bounded_render_pixels(self):
        source_path = self.images / "01.png"
        Image.new("RGB", (2000, 1000), "white").save(source_path)
        generator = self.generator()

        # The 2M-pixel source exceeds this aggregate budget, while its bounded
        # 1800x900 render (1.62M) fits. Per-source dimensions/pixels remain
        # independently limited by the production constants.
        with patch("playbook_pipeline.MAX_TOTAL_SOURCE_IMAGE_PIXELS", 1_700_000):
            image, total = generator._load_bounded_image(source_path, 0)

        self.assertEqual(image.size, (1800, 900))
        self.assertEqual(total, 1_620_000)

    def test_rejects_images_outside_print_capacity(self):
        Image.new("RGB", (32, 32), "white").save(self.images / "65.png")
        with self.assertRaisesRegex(ValueError, "Unsupported play image filename"):
            self.generator().load_images()

    def test_rejects_duplicate_numeric_aliases(self):
        Image.new("RGB", (32, 32), "white").save(self.images / "1.jpg")
        with self.assertRaisesRegex(ValueError, "Duplicate play image slot"):
            self.generator().load_images()


if __name__ == "__main__":
    unittest.main()
