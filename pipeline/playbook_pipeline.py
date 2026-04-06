#!/usr/bin/env python3
"""
Flag Football Playbook Pipeline
--------------------------------
Takes a .pptx playbook → extracts play images → generates coach cards & wristband PDFs.

Usage:
    python3 playbook_pipeline.py <playbook.pptx> [output_dir]

The script:
1. Reads the PPTX and identifies OFFENSE/DEFENSE sections
2. Converts slides to high-res images via LibreOffice + pdftoppm
3. Crops each play to its field rectangle ("Rectangle 1")
4. Names them 01.png-16.png (offense) and D1.png-D4.png (defense)
5. Feeds them into PlaybookGenerator to create coach cards + wristband PDFs
"""

import os
import sys
import subprocess
import shutil
import json
import zipfile
from pathlib import Path
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

# Import the ink overlay module
from ink_overlay import overlay_ink_on_slides

# ─── STEP 1: Analyze the PPTX ────────────────────────────────────────────────

def analyze_playbook(pptx_path):
    """
    Walk through slides, detect OFFENSE/DEFENSE sections, identify play slides.
    Returns list of dicts: {slide_index, section, play_number, crop_box_inches}
    """
    prs = Presentation(pptx_path)
    slide_width = prs.slide_width  # EMU
    slide_height = prs.slide_height

    plays = []
    current_section = None  # "OFFENSE" or "DEFENSE"
    offense_count = 0
    defense_count = 0

    for i, slide in enumerate(prs.slides):
        shapes = list(slide.shapes)
        shape_count = len(shapes)

        # Section headers have very few shapes — check for section keywords
        if shape_count <= 5:
            all_text = " ".join(
                s.text_frame.text for s in shapes if s.has_text_frame
            ).upper()
            if "OFFENSE" in all_text:
                current_section = "OFFENSE"
                offense_count = 0
                print(f"  Slide {i+1}: Section header → OFFENSE")
                continue
            elif "DEFENSE" in all_text:
                current_section = "DEFENSE"
                defense_count = 0
                print(f"  Slide {i+1}: Section header → DEFENSE")
                continue
            else:
                # Some other header/spacer — skip
                print(f"  Slide {i+1}: Skipping (header/spacer, {shape_count} shapes: {all_text[:50]})")
                continue

        if current_section is None:
            # Haven't hit a section yet — skip template slides etc.
            all_text = " ".join(
                s.text_frame.text for s in shapes if s.has_text_frame
            ).upper()
            print(f"  Slide {i+1}: Skipping (before any section, {shape_count} shapes)")
            continue

        # This is a play slide — find the main field rectangle
        crop_box = find_field_rectangle(shapes)
        if crop_box is None:
            print(f"  Slide {i+1}: Skipping (no field rectangle found)")
            continue

        # Check for special/non-play slides by looking at text
        all_text = " ".join(
            s.text_frame.text for s in shapes if s.has_text_frame
        ).upper()
        skip_keywords = ["PRINT IMAGES", "APPENDIX", "TEMPLATE"]
        if any(kw in all_text for kw in skip_keywords):
            print(f"  Slide {i+1}: Skipping (special slide: {all_text[:40]})")
            continue

        # Get play name and number from text boxes in the header area.
        # Strategy: find text boxes that START near the field top (within 5%)
        # as the header row. Ignore instruction text lower down.
        play_name = ""
        play_id = ""
        header_bottom_emu = crop_box[1]  # default to field top (no header)
        field_top = crop_box[1]
        field_height = crop_box[3] - crop_box[1]
        header_start_zone = field_top + int(field_height * 0.05)  # must start within top 5%

        header_texts = []
        for s in shapes:
            if s.has_text_frame and "TextBox" in s.name:
                text = s.text_frame.text.strip()
                if text and s.top <= header_start_zone:
                    header_texts.append((s, text))
                    tb = s.top + s.height
                    if tb > header_bottom_emu:
                        header_bottom_emu = tb

        # Separate into play_id (short/numeric) and play_name (longer text)
        for s, text in header_texts:
            if len(text) <= 3 and any(c.isdigit() for c in text):
                play_id = text
            elif len(text) <= 2 and text.isalpha():
                play_id = text  # Defense: A, B, C, D
            else:
                play_name = text

        if current_section == "OFFENSE":
            offense_count += 1
            play_num = offense_count
            filename = f"{play_num:02d}.png"
        else:
            defense_count += 1
            play_num = defense_count
            filename = f"D{play_num}.png"

        # Build label: "1 - Triple Cross" or just "D1"
        if play_id and play_name:
            label = f"{play_id} - {play_name}"
        elif play_id:
            label = play_id
        elif play_name:
            label = play_name
        else:
            label = filename.replace(".png", "")

        plays.append({
            "slide_index": i,
            "section": current_section,
            "play_number": play_num,
            "play_id": play_id,
            "play_name": play_name,
            "label": label,
            "filename": filename,
            "crop_box_emu": crop_box,
            "header_bottom_emu": header_bottom_emu,
        })
        print(f"  Slide {i+1}: {current_section} #{play_num} → {filename} ({play_id} {play_name})")

    return plays, slide_width, slide_height


def find_field_rectangle(shapes):
    """
    Find the main field rectangle on a play slide.
    Uses the LARGEST rectangle, which is the full field area.
    On end zone slides, "Rectangle 1" is the yellow end zone (too small),
    while "Rectangle 2" is the full field outline.
    Returns (left, top, right, bottom) in EMU, or None.
    """
    # Collect all rectangles with their areas
    rectangles = []
    for s in shapes:
        if hasattr(s, "shape_type") and "rectangle" in s.name.lower():
            if s.width and s.height:
                area = s.width * s.height
                rectangles.append((area, s))

    # Sort by area descending — use the largest rectangle
    rectangles.sort(key=lambda x: x[0], reverse=True)

    if rectangles:
        s = rectangles[0][1]
        return (s.left, s.top, s.left + s.width, s.top + s.height)

    # Final fallback: largest shape overall
    best = None
    best_area = 0
    for s in shapes:
        if s.width and s.height:
            area = s.width * s.height
            if area > best_area:
                best_area = area
                best = (s.left, s.top, s.left + s.width, s.top + s.height)

    return best


# ─── STEP 2: Convert PPTX to slide images ────────────────────────────────────

def convert_pptx_to_images(pptx_path, work_dir, dpi=200):
    """
    Convert PPTX → PDF → individual slide PNGs using LibreOffice + pdftoppm.
    Returns path to directory containing slide-NN.png files.
    """
    slides_dir = Path(work_dir) / "slides"
    slides_dir.mkdir(exist_ok=True)

    pdf_path = Path(work_dir) / "playbook.pdf"

    # Step 2a: PPTX → PDF via LibreOffice
    print("\n Converting PPTX → PDF via LibreOffice...")
    pptx_abs = str(Path(pptx_path).resolve())
    pptx_stem = Path(pptx_path).stem

    # Use system LibreOffice (soffice or libreoffice)
    soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_cmd:
        # macOS: LibreOffice.app doesn't add to PATH by default
        mac_soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(mac_soffice):
            soffice_cmd = mac_soffice
        else:
            raise FileNotFoundError("LibreOffice not found. Install it: brew install --cask libreoffice")
    subprocess.run([
        soffice_cmd, "--headless", "--convert-to", "pdf",
        "--outdir", str(Path(work_dir).resolve()), pptx_abs
    ], check=True, capture_output=True)

    # Find the generated PDF (could be in work_dir or next to input)
    generated_pdf = None
    for candidate in [
        Path(work_dir) / f"{pptx_stem}.pdf",
        Path(pptx_abs).with_suffix(".pdf"),
    ]:
        if candidate.exists():
            generated_pdf = candidate
            break

    if generated_pdf and str(generated_pdf) != str(pdf_path):
        shutil.move(str(generated_pdf), str(pdf_path))
    elif not pdf_path.exists():
        raise FileNotFoundError(f"LibreOffice didn't produce PDF for {pptx_abs}")

    # Step 2b: PDF → PNGs via pdftoppm
    print(f" Converting PDF → PNGs at {dpi} DPI...")
    subprocess.run([
        "pdftoppm", "-png", "-r", str(dpi),
        str(pdf_path),
        str(slides_dir / "slide")
    ], check=True, capture_output=True)

    # List generated files
    slide_images = sorted(slides_dir.glob("slide-*.png"))
    print(f" Generated {len(slide_images)} slide images")
    return slides_dir, slide_images


# ─── STEP 3: Crop play areas from slide images ───────────────────────────────

def crop_plays(plays, slide_images, slide_width_emu, slide_height_emu, output_dir):
    """
    For each play, crop the slide image to the field rectangle (including header).
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)

    if not slide_images:
        print("ERROR: No slide images found!")
        return

    # Get image dimensions from first image to compute EMU→pixel ratio
    sample = Image.open(slide_images[0])
    img_w, img_h = sample.size
    sample.close()

    emu_to_px_x = img_w / slide_width_emu
    emu_to_px_y = img_h / slide_height_emu

    for play in plays:
        si = play["slide_index"]

        # pdftoppm names slides starting from 01
        slide_num = si + 1  # 1-based
        slide_file = None
        for sf in slide_images:
            name = sf.stem  # "slide-01"
            num_str = name.split("-")[-1]
            if int(num_str) == slide_num:
                slide_file = sf
                break

        if slide_file is None:
            print(f"  WARNING: No image for slide {slide_num}, skipping {play['filename']}")
            continue

        img = Image.open(slide_file)

        # Convert crop box from EMU to pixels with small margin for routes
        left_emu, top_emu, right_emu, bottom_emu = play["crop_box_emu"]
        field_w = right_emu - left_emu
        field_h = bottom_emu - top_emu
        margin_x = int(field_w * 0.02)  # 2% horizontal margin
        margin_y = int(field_h * 0.03)  # 3% vertical margin (extra room for routes)

        left_px = max(0, int((left_emu - margin_x) * emu_to_px_x))
        top_px = max(0, int((top_emu - margin_y) * emu_to_px_y))
        right_px = min(img.width, int((right_emu + margin_x) * emu_to_px_x))
        bottom_px = min(img.height, int((bottom_emu + margin_y) * emu_to_px_y))

        cropped = img.crop((left_px, top_px, right_px, bottom_px))

        if cropped.mode != "RGB":
            cropped = cropped.convert("RGB")

        out_path = output_dir / play["filename"]
        cropped.save(out_path, "PNG")
        print(f"  {play['filename']:10s} ← Slide {slide_num} ({play['play_id']} {play['play_name']})")

    print(f"\n Saved {len(plays)} play images to {output_dir}/")


# ─── STEP 4: PlaybookGenerator (from your existing script) ───────────────────

class PlaybookGenerator:
    """Generates coach cards and wristband PDFs from play images."""

    def __init__(self, images_directory, output_directory="wristband_output"):
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.lib.units import inch
        from reportlab.lib.utils import ImageReader

        self.images_dir = Path(images_directory)
        self.output_dir = Path(output_directory)
        self.output_dir.mkdir(exist_ok=True)

    def fix_image_transparency(self, img):
        if img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            if img.mode in ('RGBA', 'LA'):
                background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else img.split()[1])
            else:
                background.paste(img)
            return background
        return img.convert('RGB') if img.mode != 'RGB' else img

    def load_images(self):
        offense_images = []
        defense_images = []
        for i in range(1, 17):
            for name in [f"{i:02d}.png", f"{i:02d}.jpg", f"{i}.png", f"{i}.jpg"]:
                img_path = self.images_dir / name
                if img_path.exists():
                    img = Image.open(img_path)
                    img = self.fix_image_transparency(img)
                    offense_images.append(img)
                    break
        for i in range(1, 7):  # Support up to D6
            for name in [f"D{i}.png", f"D{i}.jpg", f"d{i}.png", f"d{i}.jpg"]:
                img_path = self.images_dir / name
                if img_path.exists():
                    img = Image.open(img_path)
                    img = self.fix_image_transparency(img)
                    defense_images.append(img)
                    break
        return offense_images, defense_images

    def create_coach_card_offense(self, images):
        if not images:
            return
        import io
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
        from reportlab.lib.utils import ImageReader

        pdf_path = self.output_dir / "offense_coach_card.pdf"
        c = canvas.Canvas(str(pdf_path), pagesize=landscape(letter))
        page_width, page_height = landscape(letter)
        cols, rows = 4, 4
        margin = 0.5 * inch
        label_space = 0.5 * inch
        grid_width = page_width - 2 * margin - label_space
        grid_height = page_height - 2 * margin
        cell_width = grid_width / cols
        cell_height = grid_height / rows

        c.saveState()
        c.setFont("Helvetica-Bold", 24)
        c.translate(margin + label_space/2, page_height/2)
        c.rotate(90)
        c.drawCentredString(0, 0, "OFFENSE")
        c.restoreState()

        # Draw grid lines
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(0.5)
        grid_x = margin + label_space
        grid_y = margin
        for r in range(rows + 1):
            y_line = grid_y + r * cell_height
            c.line(grid_x, y_line, grid_x + grid_width, y_line)
        for cl in range(cols + 1):
            x_line = grid_x + cl * cell_width
            c.line(x_line, grid_y, x_line, grid_y + grid_height)

        for idx, img in enumerate(images[:16]):
            row = idx // cols
            col = idx % cols
            x = margin + label_space + col * cell_width
            y = page_height - (margin + (row + 1) * cell_height)
            img_buffer = io.BytesIO()
            img.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            padding = 3
            c.drawImage(ImageReader(img_buffer),
                        x + padding, y + padding,
                        width=cell_width - 2*padding,
                        height=cell_height - 2*padding,
                        preserveAspectRatio=True)
        c.save()
        print(f"  Created: {pdf_path}")

    def _defense_row_layout(self, n):
        """Return list of plays-per-row for n defense plays: 4→[2,2], 5→[2,1,2], 6→[2,2,2]."""
        if n <= 4:
            return [2, 2]
        elif n == 5:
            return [2, 1, 2]
        else:
            return [2, 2, 2]

    def create_coach_card_defense(self, images):
        if not images:
            return
        import io
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
        from reportlab.lib.utils import ImageReader

        n = len(images)
        row_layout = self._defense_row_layout(n)
        num_rows = len(row_layout)
        cols = 2

        pdf_path = self.output_dir / "defense_coach_card.pdf"
        c = canvas.Canvas(str(pdf_path), pagesize=landscape(letter))
        page_width, page_height = landscape(letter)
        margin = 1.5 * inch
        label_space = 0.5 * inch
        grid_width = page_width - 2 * margin - label_space
        grid_height = page_height - 2 * margin
        cell_width = grid_width / cols
        cell_height = grid_height / num_rows

        c.saveState()
        c.setFont("Helvetica-Bold", 24)
        c.translate(margin + label_space/2, page_height/2)
        c.rotate(90)
        c.drawCentredString(0, 0, "DEFENSE")
        c.restoreState()

        # Draw grid lines
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(0.5)
        grid_x = margin + label_space
        grid_y = margin
        for r in range(num_rows + 1):
            y_line = grid_y + r * cell_height
            c.line(grid_x, y_line, grid_x + grid_width, y_line)
        for cl in range(cols + 1):
            x_line = grid_x + cl * cell_width
            c.line(x_line, grid_y, x_line, grid_y + grid_height)

        img_idx = 0
        for row_num, count_in_row in enumerate(row_layout):
            for col_num in range(count_in_row):
                if img_idx >= n:
                    break
                if count_in_row == 1:
                    x = margin + label_space + (grid_width - cell_width) / 2
                else:
                    x = margin + label_space + col_num * cell_width
                y = page_height - (margin + (row_num + 1) * cell_height)
                img_buffer = io.BytesIO()
                images[img_idx].save(img_buffer, format='PNG')
                img_buffer.seek(0)
                padding = 10
                c.drawImage(ImageReader(img_buffer),
                            x + padding, y + padding,
                            width=cell_width - 2*padding,
                            height=cell_height - 2*padding,
                            preserveAspectRatio=True)
                img_idx += 1
        c.save()
        print(f"  Created: {pdf_path} ({n} plays, layout {'x'.join(str(r) for r in row_layout)})")

    def create_wristband_sheet_offense(self, images):
        if not images:
            return
        import io
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
        from reportlab.lib.utils import ImageReader

        pdf_path = self.output_dir / "offense_wristband.pdf"
        c = canvas.Canvas(str(pdf_path), pagesize=landscape(letter))
        page_width, page_height = landscape(letter)

        card_width = 1.0655 * inch
        card_height = 1.0205 * inch
        internal_gap = (3/64) * inch
        group_cols, group_rows = 4, 2
        groups_across, groups_down = 2, 3
        group_width = (group_cols * card_width) + ((group_cols - 1) * internal_gap)
        group_height = (group_rows * card_height) + ((group_rows - 1) * internal_gap)
        group_spacing = 0.5 * inch
        total_width = (groups_across * group_width) + ((groups_across - 1) * group_spacing)
        total_height = (groups_down * group_height) + ((groups_down - 1) * group_spacing)
        start_x = (page_width - total_width) / 2
        start_y = page_height - ((page_height - total_height) / 2)

        for page_num in range(2):
            if page_num > 0:
                c.showPage()
            start_idx = page_num * 8
            page_images = images[start_idx:start_idx + 8]
            if len(page_images) < 8:
                continue
            for group_idx in range(6):
                group_row = group_idx // groups_across
                group_col = group_idx % groups_across
                group_x = start_x + (group_col * (group_width + group_spacing))
                group_y = start_y - (group_row * (group_height + group_spacing))
                for play_idx in range(8):
                    row = play_idx // group_cols
                    col = play_idx % group_cols
                    x = group_x + (col * (card_width + internal_gap))
                    y = group_y - ((row + 1) * card_height) - (row * internal_gap)
                    c.setStrokeColorRGB(0.7, 0.7, 0.7)
                    c.setLineWidth(0.5)
                    c.rect(x, y, card_width, card_height)

                    img = page_images[play_idx]
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format='PNG')
                    img_buffer.seek(0)
                    c.drawImage(ImageReader(img_buffer),
                                x, y, width=card_width, height=card_height,
                                preserveAspectRatio=True, mask='auto')
        c.save()
        print(f"  Created: {pdf_path}")

    def create_wristband_sheet_defense(self, images):
        if not images:
            return
        import io
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
        from reportlab.lib.utils import ImageReader

        n = len(images)
        pdf_path = self.output_dir / "defense_wristband.pdf"
        c = canvas.Canvas(str(pdf_path), pagesize=landscape(letter))
        page_width, page_height = landscape(letter)

        # Same card dimensions and gap as offense wristband
        card_width = 1.0655 * inch
        card_height = 1.0205 * inch
        internal_gap = (3/64) * inch

        # Column layout: each entry = cards stacked in that column
        if n <= 4:
            col_layout = [2, 2]        # 2x2
        elif n == 5:
            col_layout = [2, 1, 2]    # 2x1x2
        else:
            col_layout = [2, 2, 2]    # 2x2x2

        num_cols = len(col_layout)

        # Match offense group dimensions (4 cols × 2 rows) for consistent cut-out size
        group_cols_ref = 4
        group_rows_ref = 2
        group_width = (group_cols_ref * card_width) + ((group_cols_ref - 1) * internal_gap)
        group_height = (group_rows_ref * card_height) + ((group_rows_ref - 1) * internal_gap)

        groups_across, groups_down = 2, 3
        group_spacing = 0.5 * inch
        total_width = (groups_across * group_width) + ((groups_across - 1) * group_spacing)
        total_height = (groups_down * group_height) + ((groups_down - 1) * group_spacing)
        start_x = (page_width - total_width) / 2
        start_y = page_height - ((page_height - total_height) / 2)

        # DEFENSE label on the left, cards centered in remaining space
        label_width = 0.25 * inch
        label_gap = 0.05 * inch
        defense_grid_width = (num_cols * card_width) + ((num_cols - 1) * internal_gap)
        cards_area = group_width - label_width - label_gap
        cards_x_offset = label_width + label_gap + (cards_area - defense_grid_width) / 2

        for group_idx in range(6):
            grow = group_idx // groups_across
            gcol = group_idx % groups_across
            group_x = start_x + (gcol * (group_width + group_spacing))
            group_y = start_y - (grow * (group_height + group_spacing))

            # Dashed cutting guide (same size as offense group)
            c.setStrokeColorRGB(0.3, 0.3, 0.3)
            c.setLineWidth(0.5)
            c.setDash([3, 3])
            c.rect(group_x, group_y - group_height, group_width, group_height)
            c.setDash([])

            # DEFENSE label (rotated 90°, left side)
            c.saveState()
            c.setFont("Helvetica-Bold", 18)
            c.translate(group_x + label_width, group_y - group_height / 2)
            c.rotate(90)
            c.drawCentredString(0, 0, "DEFENSE")
            c.restoreState()

            img_idx = 0
            for col_num, cards_in_col in enumerate(col_layout):
                for row in range(cards_in_col):
                    if img_idx >= n:
                        break
                    x = group_x + cards_x_offset + col_num * (card_width + internal_gap)
                    if cards_in_col == 1:
                        # Center single card vertically within the 2-row height
                        y = group_y - group_height / 2 - card_height / 2
                    else:
                        y = group_y - ((row + 1) * card_height) - (row * internal_gap)

                    c.setStrokeColorRGB(0.7, 0.7, 0.7)
                    c.setLineWidth(0.5)
                    c.rect(x, y, card_width, card_height)

                    img_buffer = io.BytesIO()
                    images[img_idx].save(img_buffer, format='PNG')
                    img_buffer.seek(0)
                    c.drawImage(ImageReader(img_buffer),
                                x, y, width=card_width, height=card_height,
                                preserveAspectRatio=True, mask='auto')
                    img_idx += 1

        c.save()
        print(f"  Created: {pdf_path} (6 cards, layout {'x'.join(str(r) for r in col_layout)})")

    def generate_all(self, gen_offense=True, gen_defense=True):
        print("\nLoading play images...")
        offense_images, defense_images = self.load_images()
        print(f"Found {len(offense_images)} offense plays and {len(defense_images)} defense formations")

        if gen_offense and offense_images:
            print("\nGenerating offense materials...")
            self.create_coach_card_offense(offense_images)
            self.create_wristband_sheet_offense(offense_images)
        if gen_defense and defense_images:
            print("\nGenerating defense materials...")
            self.create_coach_card_defense(defense_images)
            self.create_wristband_sheet_defense(defense_images)

        print(f"\nDone! Output in: {self.output_dir}/")
        for pdf in sorted(self.output_dir.glob("*.pdf")):
            print(f"  {pdf.name}")


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 playbook_pipeline.py <playbook.pptx> [output_dir] [--sections offense|defense|both]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else "playbook_output"

    # Parse --sections flag
    sections = "both"
    if "--sections" in sys.argv:
        idx = sys.argv.index("--sections")
        if idx + 1 < len(sys.argv):
            sections = sys.argv[idx + 1].lower()
    gen_offense = sections in ("both", "offense")
    gen_defense = sections in ("both", "defense")

    work_dir = Path("_playbook_work")
    work_dir.mkdir(exist_ok=True)
    plays_dir = work_dir / "plays"
    plays_dir.mkdir(exist_ok=True)

    print(f"{'='*60}")
    print(f"Flag Football Playbook Pipeline")
    print(f"{'='*60}")
    print(f"Input:  {pptx_path}")
    print(f"Output: {output_dir}/")
    print()

    # Step 1: Analyze
    print("STEP 1: Analyzing playbook structure...")
    plays, slide_w, slide_h = analyze_playbook(pptx_path)
    offense_plays = [p for p in plays if p["section"] == "OFFENSE"]
    defense_plays = [p for p in plays if p["section"] == "DEFENSE"]
    print(f"\n  Found {len(offense_plays)} offense plays, {len(defense_plays)} defense plays")

    # Step 2: Convert to images (skip if already done)
    slides_dir = work_dir / "slides"
    existing_slides = sorted(slides_dir.glob("slide-*.png")) if slides_dir.exists() else []
    # Only use existing slides that don't have _with_ink suffix (base slides)
    base_slides = [s for s in existing_slides if "_with_ink" not in s.name]
    if base_slides and len(base_slides) >= len(plays):
        print(f"\nSTEP 2: Using {len(base_slides)} existing slide images (skipping conversion)")
        slide_images = base_slides
    else:
        print("\nSTEP 2: Converting slides to images...")
        slides_dir, slide_images = convert_pptx_to_images(pptx_path, work_dir)

    # Step 2.5: Overlay ink annotations (hand-drawn routes)
    # Check if ink overlay was already applied (presence of _with_ink files)
    ink_files = list(slides_dir.glob("*_with_ink.png")) if slides_dir.exists() else []
    if ink_files:
        print(f"\nSTEP 2.5: Ink overlays already applied ({len(ink_files)} files), skipping...")
    else:
        print("\nSTEP 2.5: Overlaying ink annotations (hand-drawn routes)...")
        pptx_unzipped_dir = work_dir / "pptx_unzipped"
        pptx_unzipped_dir.mkdir(exist_ok=True)
        import zipfile
        with zipfile.ZipFile(pptx_path, 'r') as z:
            z.extractall(str(pptx_unzipped_dir))
        ink_output = overlay_ink_on_slides(
            pptx_path=str(Path(pptx_path).resolve()),
            pptx_unzipped_path=str(pptx_unzipped_dir),
            slides_dir=str(slides_dir),
            approach='B',
            use_fallback_if_failed=True,
            dpi=200
        )
        # Replace original slide images with ink-overlaid versions
        for slide_num, ink_path in ink_output.items():
            ink_img_path = Path(ink_path)
            for sf in slide_images:
                num_str = sf.stem.split("-")[-1]
                if int(num_str) == slide_num:
                    shutil.copy2(str(ink_img_path), str(sf))
                    print(f"  Replaced slide-{slide_num:02d}.png with ink-overlaid version")
                    break
        print(f"  Overlaid ink on {len(ink_output)} slides")

    # Step 3: Crop plays
    print("\nSTEP 3: Cropping play images...")
    crop_plays(plays, slide_images, slide_w, slide_h, plays_dir)

    # Step 4: Generate PDFs
    print(f"\nSTEP 4: Generating coach cards and wristbands (sections: {sections})...")
    generator = PlaybookGenerator(str(plays_dir), output_dir)
    generator.generate_all(gen_offense=gen_offense, gen_defense=gen_defense)

    # Cleanup
    print(f"\nPlay images saved in: {plays_dir}/")
    print(f"Final PDFs saved in: {output_dir}/")


if __name__ == "__main__":
    main()
