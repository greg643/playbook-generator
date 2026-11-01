"""Full-slide imports for Team snapshots, separate from printed-field cropping.

This is a pipeline module, not a new operator command. The job runner consumes
the returned inventory; publication still goes through the existing private
snapshot/Team authorization boundary. No guessed formation or side is emitted.
"""
import hashlib
import json
import math
import tempfile
import unicodedata
from pathlib import Path

from PIL import Image
from pptx import Presentation

from input_safety import MAX_SLIDES, validate_pptx_archive
from ink_overlay import overlay_ink_on_slides
from playbook_pipeline import convert_pptx_to_images

MAX_PAGE_BYTES = 4 * 1024 * 1024
MAX_IMPORT_BYTES = 60 * 1024 * 1024
MAX_PAGE_DIMENSION = 1800


def import_name(text, fallback):
    """Bound user-authored labels; never infer a name from arbitrary body text."""
    text = unicodedata.normalize('NFC', text or '')
    text = ''.join(' ' if unicodedata.category(c).startswith('C') else c for c in text)
    return ' '.join(text.split())[:60].strip() or fallback


def inspect_pptx_import(path):
    validate_pptx_archive(path)
    deck = Presentation(path)
    if not 1 <= len(deck.slides) <= MAX_SLIDES:
        raise ValueError(f'Import must contain 1–{MAX_SLIDES} slides')
    # A malformed custom canvas must not turn a bounded upload into a huge
    # intermediate raster, even before the per-page PNG check.
    inches = [float(deck.slide_width) / 914400, float(deck.slide_height) / 914400]
    if not all(math.isfinite(n) and 0 < n <= 40 for n in inches):
        raise ValueError('Unsupported slide dimensions')
    entries = []
    for number, slide in enumerate(deck.slides, 1):
        title = slide.shapes.title
        entries.append({'id': f'slide-{number:03d}', 'source': 'pptx', 'number': number,
                        'name': import_name(title.text if title is not None else '', f'Play {number}')})
    return entries, inches


def render_pptx_import(path, output_dir):
    """Retain every full slide in original order, including unrecognized pages.

    A private staging directory is renamed only after all images validate. A
    failed conversion never overwrites a previous ready import. The caller must
    choose a new job-scoped destination, not an account or snapshot directory.
    """
    entries, inches = inspect_pptx_import(path)
    with Path(path).open('rb') as stream:
        source_digest = hashlib.file_digest(stream, 'sha256').hexdigest()
    source_bytes = Path(path).stat().st_size
    output = Path(output_dir)
    if output.exists():
        raise ValueError('Import destination already exists')
    output.parent.mkdir(parents=True, exist_ok=True)
    dpi = min(200, int(MAX_PAGE_DIMENSION / max(inches)))
    with tempfile.TemporaryDirectory(prefix='import-pages-', dir=output.parent) as temporary:
        work = Path(temporary)
        slides_dir, images = convert_pptx_to_images(path, work, dpi=dpi)
        numbered = {}
        for image in images:
            try:
                number = int(image.stem.removeprefix('slide-'))
            except ValueError:
                raise ValueError('Invalid converted slide identity') from None
            if number in numbered:
                raise ValueError('Duplicate converted slide')
            numbered[number] = image
        if set(numbered) != set(range(1, len(entries) + 1)):
            raise ValueError('Conversion did not preserve every slide')
        overlays = overlay_ink_on_slides(str(Path(path).resolve()), '', str(slides_dir),
                                        approach='B', use_fallback_if_failed=True, dpi=dpi)
        ready = work / 'ready'
        ready.mkdir()
        total = 0
        for entry in entries:
            image = Path(overlays.get(entry['number'], numbered[entry['number']]))
            destination = ready / (entry['id'] + '.png')
            with Image.open(image) as decoded:
                width, height = decoded.size
                if (decoded.format != 'PNG' or min(width, height) < 1 or
                        max(width, height) > MAX_PAGE_DIMENSION + 1):
                    raise ValueError('Invalid imported slide dimensions')
                decoded.load()
                # Retain the entire original frame, never the field crop.
                decoded.convert('RGB').save(destination, format='PNG', optimize=True)
            size = destination.stat().st_size
            total += size
            if size > MAX_PAGE_BYTES or total > MAX_IMPORT_BYTES:
                raise ValueError('Imported playbook images exceed the snapshot budget')
            with destination.open('rb') as stream:
                digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            entry.update(file=destination.name, width=width, height=height, bytes=size, sha256=digest)
        with Path(path).open('rb') as stream:
            if hashlib.file_digest(stream, 'sha256').hexdigest() != source_digest:
                raise ValueError('Source changed during import')
        inventory = {'schema': 'gss.playbook-import-pages.v1', 'source': 'pptx',
                     'sourceSha256': source_digest, 'sourceBytes': source_bytes, 'entries': entries}
        (ready / 'import.json').write_text(json.dumps(inventory, ensure_ascii=False), encoding='utf-8')
        ready.rename(output)
    return inventory
