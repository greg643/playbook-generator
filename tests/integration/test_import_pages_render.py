"""Explicit local render gate; requires LibreOffice + Poppler (no cloud job).

Run with ~/.venv/bin/python -m unittest discover -s tests/integration
    -p test_import_pages_render.py -v
Kept out of the fast dependency-only pipeline tests; missing tools fail, not skip.
"""
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageStat

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'tests/pipeline'))
sys.path.insert(0, str(ROOT / 'pipeline'))
import test_import_pages
from import_pages import render_pptx_import


class ImportRenderTests(unittest.TestCase):
    def test_real_libreoffice_preserves_all_slides_and_names(self):
        fixture = test_import_pages.ImportPagesTests()
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        parent = ROOT / 'output/import-page-renders'
        parent.mkdir(parents=True, exist_ok=True)
        run = Path(tempfile.mkdtemp(prefix='pptx-', dir=parent))
        shutil.copy2(fixture.path, run / 'synthetic-book.pptx')
        inventory = render_pptx_import(run / 'synthetic-book.pptx', run / 'pages')
        self.assertEqual([e['name'] for e in inventory['entries']], ['Flood Left', 'DEFENSE', 'Play 3'])
        self.assertEqual([e['number'] for e in inventory['entries']], [1, 2, 3])
        for entry in inventory['entries']:
            with Image.open(run / 'pages' / entry['file']) as image:
                self.assertEqual(image.size, (1800, 900))
                if entry['number'] < 3:
                    self.assertGreater(max(ImageStat.Stat(image).stddev), 1)
        self.assertEqual(json.loads((run / 'pages/import.json').read_text()), inventory)
        print(f'Full-slide render evidence: {run}')


if __name__ == '__main__': unittest.main()
