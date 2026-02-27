#!/usr/bin/env python3
"""
Extract ink annotations from PowerPoint and overlay them onto slide images.

Supports two approaches:
- Approach A: Extract fallback PNG images from PPTX and composite them
- Approach B: Parse InkML and render strokes with PIL
"""

import os
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import io

try:
    from PIL import Image, ImageDraw
except ImportError:
    print("Pillow not installed. Installing...")
    import subprocess
    subprocess.check_call(['pip', 'install', 'Pillow'])
    from PIL import Image, ImageDraw

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except ImportError:
    print("python-pptx not installed. Installing...")
    import subprocess
    subprocess.check_call(['pip', 'install', 'python-pptx'])
    from pptx import Presentation
    from pptx.util import Inches, Pt


# XML namespaces
NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'inkml': 'http://www.w3.org/2003/InkML',
}


def emu_to_pixels(emu_value: float, dpi: int = 96, emu_per_inch: int = 914400) -> float:
    """Convert EMU (English Metric Units) to pixels."""
    inches = emu_value / emu_per_inch
    return inches * dpi


def load_pptx_zip(pptx_path: str) -> zipfile.ZipFile:
    """Load PPTX as a ZIP file."""
    return zipfile.ZipFile(pptx_path, 'r')


def extract_slide_relationships(pptx_zip: zipfile.ZipFile, slide_num: int) -> Dict[str, str]:
    """Extract relationship mappings for a slide, resolving relative paths."""
    from pathlib import PurePath

    rel_path = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
    try:
        rel_content = pptx_zip.read(rel_path).decode('utf-8')
        root = ET.fromstring(rel_content)
        relationships = {}
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rel_id = rel.get('Id')
            target = rel.get('Target')

            # Resolve relative paths
            if target and target.startswith('..'):
                # Resolve relative to ppt/slides
                resolved = str(PurePath('ppt/slides') / target).replace('\\', '/')
                # Normalize the path
                parts = resolved.split('/')
                normalized = []
                for part in parts:
                    if part == '..':
                        if normalized:
                            normalized.pop()
                    elif part != '.':
                        normalized.append(part)
                target = '/'.join(normalized)

            relationships[rel_id] = target
        return relationships
    except KeyError:
        return {}


def extract_fallback_images(pptx_unzipped_path: str, pptx_path: str, slide_num: int) -> List[Tuple[str, bytes, Tuple[int, int, int, int]]]:
    """
    Extract fallback PNG images from slide XML.
    Returns list of (rel_id, image_data, position_tuple) where position is (x_emu, y_emu, width_emu, height_emu)
    """
    pptx_zip = load_pptx_zip(pptx_path)
    slide_path = f'ppt/slides/slide{slide_num}.xml'

    try:
        slide_content = pptx_zip.read(slide_path).decode('utf-8')
    except KeyError:
        pptx_zip.close()
        return []

    root = ET.fromstring(slide_content)
    fallback_images = []

    # Get slide relationships
    relationships = extract_slide_relationships(pptx_zip, slide_num)

    # Find all mc:Fallback elements (search all, not just within grpSp)
    for fallback in root.findall('.//{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback'):
        # Look for picture inside fallback
        pic = fallback.find(f'{{{NS["p"]}}}pic')
        if pic is None:
            continue

        # Get the blip reference
        blip = pic.find(f'.//{{{NS["a"]}}}blip')
        if blip is None:
            continue

        rel_id = blip.get(f'{{{NS["r"]}}}embed')
        if not rel_id or rel_id not in relationships:
            continue

        # Get the position from the picture shape properties
        spPr = pic.find(f'{{{NS["p"]}}}spPr')
        if spPr is None:
            continue

        xfrm = spPr.find(f'{{{NS["a"]}}}xfrm')
        if xfrm is None:
            continue

        off = xfrm.find(f'{{{NS["a"]}}}off')
        ext = xfrm.find(f'{{{NS["a"]}}}ext')

        if off is None or ext is None:
            continue

        x = int(off.get('x', 0))
        y = int(off.get('y', 0))
        w = int(ext.get('cx', 0))
        h = int(ext.get('cy', 0))

        # Get actual image file
        image_target = relationships[rel_id]
        # image_target is already a full path like "ppt/media/image312.png"
        try:
            image_data = pptx_zip.read(image_target)
            fallback_images.append((rel_id, image_data, (x, y, w, h)))
        except KeyError:
            pass

    pptx_zip.close()
    return fallback_images


def extract_ink_strokes(pptx_unzipped_path: str, slide_num: int, pptx_zip: zipfile.ZipFile) -> List[Tuple[List[Tuple[float, float]], Tuple[int, int, int, int], str]]:
    """
    Extract ink strokes from InkML files with position and color information.

    Handles contentParts that are:
    1. At the slide root level
    2. Inside groups (p:grpSp) with coordinate transforms

    Returns list of (stroke_points, position_tuple, color_hex) where:
    - stroke_points: list of (x, y) absolute coordinates in 1/1000cm units
    - position_tuple: (x_emu, y_emu, width_emu, height_emu) - the bounding box of the ink annotation on the slide
    - color_hex: hex color string (e.g., '#FF0000', '#000000')
    """
    slide_path = f'ppt/slides/slide{slide_num}.xml'

    try:
        slide_content = pptx_zip.read(slide_path).decode('utf-8')
    except KeyError:
        return []

    root = ET.fromstring(slide_content)
    ink_strokes = []

    # Get relationships once
    relationships = extract_slide_relationships(pptx_zip, slide_num)

    # Build a map of all contentParts with their parent group transform (if any)
    # Structure: {content_part_element: group_transform_dict_or_None}
    content_part_transforms = {}

    # First, find top-level contentParts (not inside groups)
    # They can be at the root level or inside mc:AlternateContent/mc:Choice
    # Note: The elements are p:contentPart (with 'p' namespace), not p14:contentPart
    top_level_parts = root.findall('.//p:cSld/p:spTree/mc:AlternateContent/mc:Choice/p:contentPart', NS)
    for cp in top_level_parts:
        content_part_transforms[cp] = None  # No group transform

    # Also find direct contentPart elements
    top_level_parts2 = root.findall('.//p:cSld/p:spTree/p:contentPart', NS)
    for cp in top_level_parts2:
        content_part_transforms[cp] = None

    # Now find contentParts inside groups
    # Iterate through all groups on the slide
    for group in root.findall('.//p:cSld/p:spTree/p:grpSp', NS):
        group_xfrm = group.find(f'{{{NS["p"]}}}grpSpPr/{{{NS["a"]}}}xfrm', NS)

        if group_xfrm is None:
            continue

        # Get group's position on slide (a:off, a:ext)
        group_off = group_xfrm.find(f'{{{NS["a"]}}}off', NS)
        group_ext = group_xfrm.find(f'{{{NS["a"]}}}ext', NS)
        # Get group's child coordinate space (a:chOff, a:chExt)
        group_chOff = group_xfrm.find(f'{{{NS["a"]}}}chOff', NS)
        group_chExt = group_xfrm.find(f'{{{NS["a"]}}}chExt', NS)

        if group_off is None or group_ext is None:
            continue

        group_x = int(group_off.get('x', 0))
        group_y = int(group_off.get('y', 0))
        group_w = int(group_ext.get('cx', 1))  # Avoid division by zero
        group_h = int(group_ext.get('cy', 1))

        # Child coordinate space offsets (default to 0,0 if not specified)
        choff_x = int(group_chOff.get('x', 0)) if group_chOff is not None else 0
        choff_y = int(group_chOff.get('y', 0)) if group_chOff is not None else 0
        chext_cx = int(group_chExt.get('cx', group_w)) if group_chExt is not None else group_w
        chext_cy = int(group_chExt.get('cy', group_h)) if group_chExt is not None else group_h

        # Find all contentParts inside this group
        # The elements are p:contentPart, nested inside mc:AlternateContent/mc:Choice
        group_content_parts = group.findall('.//mc:Choice/p:contentPart', NS)

        for cp in group_content_parts:
            # Store the group transform info with the content part
            content_part_transforms[cp] = {
                'group_x': group_x,
                'group_y': group_y,
                'group_w': group_w,
                'group_h': group_h,
                'choff_x': choff_x,
                'choff_y': choff_y,
                'chext_cx': chext_cx,
                'chext_cy': chext_cy
            }

    # Now process all contentParts
    for content_part, group_transform in content_part_transforms.items():
        # Get the rel_id pointing to the ink file
        # The r:id attribute might be on the contentPart itself
        rel_id = content_part.get(f'{{{NS["r"]}}}id')

        # If not found, look in child elements (sometimes it's in a different location)
        if not rel_id:
            # Check if there's an AlternateContent/Choice wrapper with the r:id
            continue

        # Get position from the xfrm element (the ink's bounding box)
        # This position is in the coordinate space (child space if inside group, slide space if top-level)
        # The xfrm element is inside a p14:nvContentPartPr or similar structure
        part_xfrm = content_part.find(f'{{{NS["p14"]}}}xfrm', NS)

        # If not found with p14 namespace, try without namespace prefix (it's still in the element tree)
        if part_xfrm is None:
            # Try finding it with broader search
            for elem in content_part.iter():
                if 'xfrm' in elem.tag:
                    part_xfrm = elem
                    break

        if part_xfrm is None:
            continue

        off = part_xfrm.find(f'{{{NS["a"]}}}off', NS)
        ext = part_xfrm.find(f'{{{NS["a"]}}}ext', NS)

        if off is None or ext is None:
            continue

        x = int(off.get('x', 0))
        y = int(off.get('y', 0))
        w = int(ext.get('cx', 0))
        h = int(ext.get('cy', 0))

        # If the contentPart is inside a group, transform its coordinates from child space to slide space
        if group_transform is not None:
            # Apply group transform: slide_coord = group_offset + (child_coord - chOff) * (group_ext / chExt)
            scale_x = group_transform['group_w'] / group_transform['chext_cx'] if group_transform['chext_cx'] != 0 else 1
            scale_y = group_transform['group_h'] / group_transform['chext_cy'] if group_transform['chext_cy'] != 0 else 1

            x = group_transform['group_x'] + int((x - group_transform['choff_x']) * scale_x)
            y = group_transform['group_y'] + int((y - group_transform['choff_y']) * scale_y)
            w = int(w * scale_x)
            h = int(h * scale_y)

        # Get the ink file
        if rel_id not in relationships:
            continue

        ink_target = relationships[rel_id]
        # ink_target already includes the full path like "ppt/ink/ink9.xml"
        ink_path = ink_target if ink_target.startswith('ppt/') else f'ppt/{ink_target}'

        try:
            ink_content = pptx_zip.read(ink_path).decode('utf-8')
            strokes = parse_inkml(ink_content)
            # strokes is now list of (points, dummy_pos, color)
            for stroke_points, _, color_hex in strokes:
                ink_strokes.append((stroke_points, (x, y, w, h), color_hex))
        except KeyError:
            pass

    return ink_strokes


def parse_inkml(inkml_content: str) -> List[Tuple[List[Tuple[float, float]], Tuple[int, int, int, int], str]]:
    """
    Parse InkML XML and extract stroke coordinates with color info.

    Returns list of (stroke_points, (0,0,0,0), color_hex).

    InkML delta encoding:
    - First point is absolute: "3055 4785 24575"
    - ' prefix = first-order delta (velocity): delta = value, pos += delta
    - " prefix = second-order delta (accel): delta += value, pos += delta
    - No prefix after deltas = continues previous mode
    - Segments separated by commas; values within segment use regex extraction
      because minus signs double as separators (e.g. "4-2 0" = X=4, Y=-2, F=0)
    """
    import re

    root = ET.fromstring(inkml_content)
    strokes = []
    brush_color = '#000000'

    for trace in root.findall('.//inkml:trace', NS):
        trace_text = trace.text
        if not trace_text:
            continue

        # Look up brush color
        brush_ref = trace.get('brushRef')
        trace_brush_color = brush_color
        if brush_ref:
            brush_id = brush_ref.lstrip('#')
            for brush in root.findall('.//inkml:brush', NS):
                if brush.get('{http://www.w3.org/XML/1998/namespace}id') == brush_id:
                    for prop in brush.findall('inkml:brushProperty', NS):
                        if prop.get('name') == 'color':
                            trace_brush_color = prop.get('value')
                            break

        points = []
        points = []
        pos_x, pos_y = 0.0, 0.0
        vel_x, vel_y = 0.0, 0.0  # velocity (first-order delta)
        mode = 'abs'  # 'abs', 'vel' (first-order), 'acc' (second-order)

        segments = trace_text.split(',')
        for i, seg in enumerate(segments):
            seg = seg.strip()
            if not seg:
                continue
            try:
                # Use regex for ALL segments — handles both "4-2 0" and "'0'-9'0"
                matches = list(re.finditer(r"""(["']?)(-?\d+)""", seg))
                if len(matches) < 2:
                    continue

                x_quote = matches[0].group(1)
                x_val = float(matches[0].group(2))
                y_quote = matches[1].group(1)
                y_val = float(matches[1].group(2))

                if i == 0:
                    # First segment: absolute position
                    pos_x, pos_y = x_val, y_val
                    vel_x, vel_y = 0.0, 0.0
                    mode = 'abs'
                else:
                    # Update mode based on quote prefixes (only if explicitly marked)
                    if x_quote == '"' or y_quote == '"':
                        mode = 'acc'
                    elif x_quote == "'" or y_quote == "'":
                        mode = 'vel'
                    # else: no quotes → continue in previous mode

                    if mode == 'acc':
                        # Second-order: values are acceleration
                        vel_x += x_val
                        vel_y += y_val
                    else:
                        # First-order: values are velocity (direct delta)
                        vel_x = x_val
                        vel_y = y_val

                    pos_x += vel_x
                    pos_y += vel_y

                points.append((pos_x, pos_y))
            except (ValueError, IndexError):
                continue

        if points:
            strokes.append((points, (0, 0, 0, 0), trace_brush_color))

    return strokes


def get_slide_dimensions(pptx_path: str) -> Tuple[int, int]:
    """Get slide dimensions in EMU from PPTX."""
    prs = Presentation(pptx_path)
    # slide width and height in EMU
    return (prs.slide_width, prs.slide_height)


def overlay_fallback_images_approach_a(
    slide_image_path: str,
    fallback_images: List[Tuple[str, bytes, Tuple[int, int, int, int]]],
    slide_width_emu: int,
    slide_height_emu: int,
    slide_image_width_px: int,
    slide_image_height_px: int,
    dpi: int = 96
) -> Image.Image:
    """
    Approach A: Composite fallback PNG images onto the slide image.
    """
    # Load the slide image
    slide_image = Image.open(slide_image_path).convert('RGBA')

    # Calculate scaling factor from EMU to pixels
    emu_per_inch = 914400
    slide_width_inches = slide_width_emu / emu_per_inch
    slide_width_expected_px = slide_width_inches * dpi
    scale_factor = slide_image_width_px / slide_width_expected_px

    for rel_id, image_data, (x_emu, y_emu, w_emu, h_emu) in fallback_images:
        try:
            # Load fallback image
            fallback_img = Image.open(io.BytesIO(image_data)).convert('RGBA')

            # Convert EMU to pixels
            x_px = int(emu_to_pixels(x_emu, dpi) * scale_factor)
            y_px = int(emu_to_pixels(y_emu, dpi) * scale_factor)
            w_px = int(emu_to_pixels(w_emu, dpi) * scale_factor)
            h_px = int(emu_to_pixels(h_emu, dpi) * scale_factor)

            # Resize fallback image to match position size
            fallback_resized = fallback_img.resize((w_px, h_px), Image.Resampling.LANCZOS)

            # Paste onto slide image
            slide_image.paste(fallback_resized, (x_px, y_px), fallback_resized)
        except Exception as e:
            print(f"Error processing fallback image {rel_id}: {e}")
            continue

    return slide_image


def hex_to_rgba(hex_color: str, alpha: int = 255) -> Tuple[int, int, int, int]:
    """Convert hex color string (e.g., '#FF0000') to RGBA tuple."""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b, alpha)
    return (0, 0, 0, alpha)  # Default to black


def overlay_inkml_strokes_approach_b(
    slide_image_path: str,
    ink_strokes: List[Tuple[List[Tuple[float, float]], Tuple[int, int, int, int], str]],
    slide_width_emu: int,
    slide_height_emu: int,
    slide_image_width_px: int,
    slide_image_height_px: int,
    dpi: int = 96,
    stroke_width: int = 2
) -> Image.Image:
    """
    Approach B: Parse InkML and render strokes onto the slide image.

    The InkML coordinates are in 1/1000cm units (10 = 0.1mm, 1000 = 1cm).
    The strokes have absolute coordinates from the InkML file.
    We need to map them from their coordinate space to the bounding box position on the slide.

    The mapping works as follows:
    1. InkML coordinates are absolute in 1/1000cm units
    2. The bounding box (x_emu, y_emu, w_emu, h_emu) defines where the ink should appear on the slide
    3. We normalize InkML coordinates by their min/max, then scale to the bounding box size and position
    """
    # Load the slide image
    slide_image = Image.open(slide_image_path).convert('RGBA')
    draw = ImageDraw.Draw(slide_image)

    # Calculate scaling factor from EMU to pixels at the given DPI
    emu_per_inch = 914400
    slide_width_inches = slide_width_emu / emu_per_inch
    slide_width_expected_px = slide_width_inches * dpi
    scale_factor = slide_image_width_px / slide_width_expected_px

    for stroke_points, (x_emu, y_emu, w_emu, h_emu), color_hex in ink_strokes:
        if len(stroke_points) < 2:
            continue

        # Convert bounding box from EMU to pixels
        x_px = emu_to_pixels(x_emu, dpi) * scale_factor
        y_px = emu_to_pixels(y_emu, dpi) * scale_factor
        w_px = emu_to_pixels(w_emu, dpi) * scale_factor
        h_px = emu_to_pixels(h_emu, dpi) * scale_factor

        # Find the bounding box of the stroke points in InkML coordinate space
        # InkML uses 1/1000cm units
        min_x = min(p[0] for p in stroke_points)
        max_x = max(p[0] for p in stroke_points)
        min_y = min(p[1] for p in stroke_points)
        max_y = max(p[1] for p in stroke_points)

        # Calculate the range (handle case where all points are the same)
        range_x = max_x - min_x if max_x > min_x else 1
        range_y = max_y - min_y if max_y > min_y else 1

        # Convert color from hex to RGBA
        stroke_color = hex_to_rgba(color_hex)

        # Estimate stroke width in pixels from InkML brush width
        # InkML brush width is in cm (0.05 cm), convert to pixels
        brush_width_px = emu_to_pixels(0.05 * 360000, dpi)  # 0.05cm in EMU
        brush_width_px = max(1, int(brush_width_px * scale_factor))

        # Draw the stroke
        for i in range(len(stroke_points) - 1):
            p1 = stroke_points[i]
            p2 = stroke_points[i + 1]

            # Normalize InkML coordinates to 0-1 range based on the stroke's bounding box
            norm_x1 = (p1[0] - min_x) / range_x
            norm_y1 = (p1[1] - min_y) / range_y
            norm_x2 = (p2[0] - min_x) / range_x
            norm_y2 = (p2[1] - min_y) / range_y

            # Convert normalized coordinates to pixel coordinates in the slide image
            px1 = x_px + norm_x1 * w_px
            py1 = y_px + norm_y1 * h_px
            px2 = x_px + norm_x2 * w_px
            py2 = y_px + norm_y2 * h_px

            draw.line([(px1, py1), (px2, py2)], fill=stroke_color, width=brush_width_px)

    return slide_image


def overlay_ink_on_slides(
    pptx_path: str,
    pptx_unzipped_path: str,
    slides_dir: str,
    approach: str = 'B',
    use_fallback_if_failed: bool = True,
    dpi: int = 96
) -> Dict[int, str]:
    """
    Main function to overlay ink annotations on slide images.

    Args:
        pptx_path: Path to the PPTX file
        pptx_unzipped_path: Path to unzipped PPTX directory
        slides_dir: Directory containing slide PNG images (e.g., slide-01.png, slide-02.png)
        approach: 'B' for InkML parsing (primary), 'A' for fallback images
        use_fallback_if_failed: If approach B fails, fall back to approach A
        dpi: DPI for conversion

    Returns:
        Dictionary mapping slide numbers to output image paths
    """
    pptx_zip = load_pptx_zip(pptx_path)

    # Get slide dimensions
    slide_width_emu, slide_height_emu = get_slide_dimensions(pptx_path)

    # Find all slide images
    slides_path = Path(slides_dir)
    slide_files = sorted(slides_path.glob('slide-*.png'))

    output_paths = {}

    for slide_file in slide_files:
        # Extract slide number
        try:
            slide_num = int(slide_file.stem.split('-')[1])
        except (ValueError, IndexError):
            continue

        print(f"Processing slide {slide_num}...")

        # Get actual image dimensions
        with Image.open(str(slide_file)) as img:
            slide_image_width_px, slide_image_height_px = img.size

        result_image = None

        if approach == 'B':
            # Approach B: InkML parsing (primary approach)
            try:
                ink_strokes = extract_ink_strokes(
                    pptx_unzipped_path,
                    slide_num,
                    pptx_zip
                )

                if ink_strokes:
                    result_image = overlay_inkml_strokes_approach_b(
                        str(slide_file),
                        ink_strokes,
                        slide_width_emu,
                        slide_height_emu,
                        slide_image_width_px,
                        slide_image_height_px,
                        dpi
                    )
                    print(f"  - Approach B (InkML): Rendered {len(ink_strokes)} ink strokes")
            except Exception as e:
                print(f"  - Approach B (InkML) failed: {e}")
                import traceback
                traceback.print_exc()

        elif approach == 'A':
            # Approach A: Fallback images
            try:
                fallback_images = extract_fallback_images(
                    pptx_unzipped_path,
                    pptx_path,
                    slide_num
                )

                if fallback_images:
                    result_image = overlay_fallback_images_approach_a(
                        str(slide_file),
                        fallback_images,
                        slide_width_emu,
                        slide_height_emu,
                        slide_image_width_px,
                        slide_image_height_px,
                        dpi
                    )
                    print(f"  - Approach A (Fallback): Composited {len(fallback_images)} fallback images")
            except Exception as e:
                print(f"  - Approach A (Fallback) failed: {e}")

        # Fall back to other approach if requested
        if result_image is None and use_fallback_if_failed:
            fallback_approach = 'A' if approach == 'B' else 'B'
            print(f"  - Falling back to Approach {fallback_approach}...")

            try:
                if fallback_approach == 'A':
                    fallback_images = extract_fallback_images(
                        pptx_unzipped_path,
                        pptx_path,
                        slide_num
                    )
                    if fallback_images:
                        result_image = overlay_fallback_images_approach_a(
                            str(slide_file),
                            fallback_images,
                            slide_width_emu,
                            slide_height_emu,
                            slide_image_width_px,
                            slide_image_height_px,
                            dpi
                        )
                        print(f"    - Fallback Approach A (Fallback): Composited {len(fallback_images)} fallback images")
                else:
                    ink_strokes = extract_ink_strokes(
                        pptx_unzipped_path,
                        slide_num,
                        pptx_zip
                    )
                    if ink_strokes:
                        result_image = overlay_inkml_strokes_approach_b(
                            str(slide_file),
                            ink_strokes,
                            slide_width_emu,
                            slide_height_emu,
                            slide_image_width_px,
                            slide_image_height_px,
                            dpi
                        )
                        print(f"    - Fallback Approach B (InkML): Rendered {len(ink_strokes)} ink strokes")
            except Exception as e:
                print(f"    - Fallback approach also failed: {e}")

        # If we have a result, save it
        if result_image is not None:
            output_path = slide_file.parent / f"slide-{slide_num:02d}_with_ink.png"
            result_image.save(str(output_path), 'PNG')
            output_paths[slide_num] = str(output_path)
            print(f"  - Saved to {output_path}")
        else:
            print(f"  - No ink annotations found or all approaches failed")

    pptx_zip.close()
    return output_paths


if __name__ == '__main__':
    import sys
    import tempfile

    if len(sys.argv) < 2:
        print("Usage: python ink_overlay.py <playbook.pptx> [slides_dir] [approach]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    slides_dir = sys.argv[2] if len(sys.argv) > 2 else './slides'
    approach = sys.argv[3] if len(sys.argv) > 3 else 'B'

    # Unzip PPTX to a temp directory
    pptx_unzipped_path = tempfile.mkdtemp(prefix="pptx_unzipped_")
    import zipfile as _zf
    with _zf.ZipFile(pptx_path, 'r') as z:
        z.extractall(pptx_unzipped_path)

    print(f"Using approach {approach}")
    output_paths = overlay_ink_on_slides(
        pptx_path,
        pptx_unzipped_path,
        slides_dir,
        approach=approach
    )

    print(f"\nProcessed {len(output_paths)} slides")
    for slide_num, output_path in sorted(output_paths.items()):
        print(f"  Slide {slide_num}: {output_path}")
