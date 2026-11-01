import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'pipeline'))
import import_pages


class ImportPagesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.path = self.root / 'book.pptx'
        self.output = self.root / 'import'
        deck = Presentation()
        deck.slide_width, deck.slide_height = Inches(10), Inches(5)
        slide = deck.slides.add_slide(deck.slide_layouts[0])
        slide.shapes.title.text = '  Flood\nLeft  '
        # No field rectangle, header and untitled pages must not disappear.
        slide = deck.slides.add_slide(deck.slide_layouts[0])
        slide.shapes.title.text = 'DEFENSE'
        deck.slides.add_slide(deck.slide_layouts[6])
        deck.save(self.path)

    def convert(self, path, work, dpi):
        self.assertEqual(dpi, 180)
        slides = Path(work) / 'slides'; slides.mkdir()
        for i in range(1, 4):
            Image.new('RGB', (1800, 900), (i * 50, 10, 20)).save(slides / f'slide-{i}.png')
        # Deliberately scrambled input demonstrates identity/order, not glob order.
        return slides, [slides / f'slide-{i}.png' for i in (3, 1, 2)]

    def render(self, converter=None, overlays=None):
        with patch.object(import_pages, 'convert_pptx_to_images', side_effect=converter or self.convert), \
             patch.object(import_pages, 'overlay_ink_on_slides', return_value=overlays or {}):
            return import_pages.render_pptx_import(self.path, self.output)

    def test_inventory_keeps_every_slide_title_and_original_order(self):
        entries, size = import_pages.inspect_pptx_import(self.path)
        self.assertEqual(size, [10, 5])
        self.assertEqual([e['name'] for e in entries], ['Flood Left', 'DEFENSE', 'Play 3'])
        self.assertEqual([e['number'] for e in entries], [1, 2, 3])
        self.assertTrue(all(e['source'] == 'pptx' for e in entries))
        self.assertTrue(all('playersPerSide' not in e and 'side' not in e for e in entries))

    def test_full_frames_not_cropped_and_digest_matches_retained_bytes(self):
        inventory = self.render()
        self.assertEqual(inventory['schema'], 'gss.playbook-import-pages.v1')
        self.assertEqual(json.loads((self.output / 'import.json').read_text()), inventory)
        self.assertEqual(inventory['sourceSha256'], hashlib.sha256(self.path.read_bytes()).hexdigest())
        for i, entry in enumerate(inventory['entries'], 1):
            path = self.output / entry['file']
            self.assertEqual((entry['width'], entry['height']), (1800, 900))
            self.assertEqual(entry['bytes'], path.stat().st_size)
            self.assertEqual(entry['sha256'], hashlib.sha256(path.read_bytes()).hexdigest())
            with Image.open(path) as image:
                self.assertEqual(image.getpixel((0, 0)), (i * 50, 10, 20))
                self.assertEqual(image.getpixel((1799, 899)), (i * 50, 10, 20))

    def test_missing_or_duplicate_page_never_publishes_partial_import(self):
        for mode in ('missing', 'duplicate'):
            def broken(path, work, dpi):
                folder, images = self.convert(path, work, dpi)
                return folder, images[:-1] if mode == 'missing' else images + images[:1]
            with self.subTest(mode=mode), self.assertRaises(ValueError): self.render(broken)
            self.assertFalse(self.output.exists())

    def test_existing_ready_import_is_never_overwritten(self):
        self.output.mkdir(); marker = self.output / 'keep.txt'; marker.write_text('current')
        with self.assertRaisesRegex(ValueError, 'already exists'): self.render()
        self.assertEqual(marker.read_text(), 'current')

    def test_budget_or_invalid_frame_never_leaves_ready_output(self):
        with patch.object(import_pages, 'MAX_IMPORT_BYTES', 10):
            with self.assertRaisesRegex(ValueError, 'budget'): self.render()
        self.assertFalse(self.output.exists())
        overlay = self.root / 'too-large.png'
        Image.new('RGB', (1900, 900)).save(overlay)
        with self.assertRaisesRegex(ValueError, 'dimensions'): self.render(overlays={1: overlay})
        self.assertFalse(self.output.exists())

    def test_ink_overlay_replaces_only_its_bound_slide(self):
        overlay = self.root / 'ink.png'; Image.new('RGB', (1800, 900), 'green').save(overlay)
        inventory = self.render(overlays={2: overlay})
        with Image.open(self.output / inventory['entries'][1]['file']) as image:
            self.assertEqual(image.getpixel((1, 1)), (0, 128, 0))
        with Image.open(self.output / inventory['entries'][0]['file']) as image:
            self.assertEqual(image.getpixel((1, 1)), (50, 10, 20))

    def test_empty_and_oversized_decks_fail_before_conversion(self):
        deck = Presentation(); deck.save(self.path)
        with self.assertRaisesRegex(ValueError, '1–100'): import_pages.inspect_pptx_import(self.path)
        deck.slides.add_slide(deck.slide_layouts[6]); deck.slide_width = Inches(41); deck.save(self.path)
        with self.assertRaisesRegex(ValueError, 'dimensions'): import_pages.inspect_pptx_import(self.path)

    def test_names_are_bounded_normalized_and_controls_removed(self):
        self.assertEqual(import_pages.import_name('Cafe\u0301\tGo\x00Now', 'Play 1'), 'Café Go Now')
        self.assertEqual(import_pages.import_name('\n\t', 'Play 1'), 'Play 1')
        self.assertEqual(len(import_pages.import_name('x' * 100, 'Play 1')), 60)

    def test_source_change_during_conversion_never_commits(self):
        def changed(path, work, dpi):
            result = self.convert(path, work, dpi)
            with self.path.open('ab') as stream: stream.write(b'changed')
            return result
        with self.assertRaisesRegex(ValueError, 'Source changed'): self.render(changed)
        self.assertFalse(self.output.exists())

    def test_existing_command_routes_imports_without_field_detection(self):
        import playbook_pipeline
        result = {'entries': [{'id': 'slide-001'}]}
        with patch.object(sys, 'argv', ['playbook_pipeline.py', str(self.path), str(self.output), '--import-pages']), \
             patch.object(import_pages, 'render_pptx_import', return_value=result) as render, \
             patch.object(playbook_pipeline, 'analyze_playbook') as crop:
            self.assertEqual(playbook_pipeline.main(), result)
            render.assert_called_once_with(str(self.path), str(self.output))
            crop.assert_not_called()


if __name__ == '__main__': unittest.main()
