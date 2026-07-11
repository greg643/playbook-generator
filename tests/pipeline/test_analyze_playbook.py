import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from playbook_pipeline import analyze_playbook  # noqa: E402


class AnalyzePlaybookTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.path = Path(self.temp_dir.name) / "deck.pptx"

    @staticmethod
    def add_title(slide, text):
        box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(3), Inches(0.5))
        box.text_frame.text = text

    @staticmethod
    def add_play(slide, name):
        slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(1),
            Inches(1),
            Inches(8),
            Inches(5),
        )
        box = slide.shapes.add_textbox(Inches(2), Inches(1.05), Inches(3), Inches(0.4))
        box.text_frame.text = name

    def test_decorated_headers_simple_plays_and_repeated_sections(self):
        presentation = Presentation()
        presentation.slides._sldIdLst.clear()
        blank = presentation.slide_layouts[6]

        header = presentation.slides.add_slide(blank)
        self.add_title(header, "OFFENSE")
        for index in range(6):
            header.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(index), Inches(2), Inches(0.2), Inches(0.2),
            )

        play_one = presentation.slides.add_slide(blank)
        self.add_play(play_one, "First")

        repeated_header = presentation.slides.add_slide(blank)
        self.add_title(repeated_header, "OFFENSE")
        play_two = presentation.slides.add_slide(blank)
        self.add_play(play_two, "Second")

        no_field = presentation.slides.add_slide(blank)
        no_field.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(8), Inches(5))

        defense_header = presentation.slides.add_slide(blank)
        self.add_title(defense_header, "DEFENSE")
        defense_play = presentation.slides.add_slide(blank)
        self.add_play(defense_play, "Cover")

        presentation.save(self.path)
        plays, _width, _height = analyze_playbook(self.path)

        self.assertEqual([play["filename"] for play in plays], ["01.png", "02.png", "D1.png"])
        self.assertEqual([play["section"] for play in plays], ["OFFENSE", "OFFENSE", "DEFENSE"])

    def test_free_text_case_insensitive_headers(self):
        # Real decks decorate their separators: Greg's deck says "6v6 OFFENSE".
        # A play slide that merely mentions a section (name "Offense Smash",
        # field rectangle present) must stay a play, not become a header.
        presentation = Presentation()
        presentation.slides._sldIdLst.clear()
        blank = presentation.slide_layouts[6]

        offense_header = presentation.slides.add_slide(blank)
        self.add_title(offense_header, "6v6 Offense")
        named_play = presentation.slides.add_slide(blank)
        self.add_play(named_play, "Offense Smash")

        defense_header = presentation.slides.add_slide(blank)
        self.add_title(defense_header, "DEFENSE — Zone Looks")
        defense_play = presentation.slides.add_slide(blank)
        self.add_play(defense_play, "Cover 2")

        presentation.save(self.path)
        plays, _width, _height = analyze_playbook(self.path)

        self.assertEqual([play["filename"] for play in plays], ["01.png", "D1.png"])
        self.assertEqual(plays[0]["play_name"], "Offense Smash")

    def test_no_sections_raises_guideline(self):
        presentation = Presentation()
        presentation.slides._sldIdLst.clear()
        play_only = presentation.slides.add_slide(presentation.slide_layouts[6])
        self.add_play(play_only, "Orphan")
        presentation.save(self.path)

        with self.assertRaisesRegex(ValueError, "separator slide"):
            analyze_playbook(self.path)

    def test_sections_without_plays_raises_guideline(self):
        presentation = Presentation()
        presentation.slides._sldIdLst.clear()
        header = presentation.slides.add_slide(presentation.slide_layouts[6])
        self.add_title(header, "Offense")
        presentation.save(self.path)

        with self.assertRaisesRegex(ValueError, "no play slides with a field rectangle"):
            analyze_playbook(self.path)


if __name__ == "__main__":
    unittest.main()
