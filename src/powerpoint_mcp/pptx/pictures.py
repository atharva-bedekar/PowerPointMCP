"""Core operations for PowerPoint pictures: insertion, replacement, and geometry calculation."""

from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union
from PIL import Image

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

from powerpoint_mcp.models.shape import (
    emu_to_inches,
    inches_to_emu,
)


def _get_image_metadata(image_path: Union[str, Path]) -> Tuple[int, int, str, float]:
    """Extract width, height, format, and aspect ratio from an image file.

    Returns:
        (pixel_width, pixel_height, format_name, aspect_ratio_w_over_h)
    """
    path = Path(image_path)
    if not path.is_file():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    with Image.open(path) as img:
        w, h = img.size
        img_format = img.format or "PNG"
        aspect = float(w) / float(h) if h > 0 else 1.0
        return w, h, img_format, aspect


def calculate_picture_dimensions(
    px_width: int,
    px_height: int,
    requested_width: Optional[float] = None,
    requested_height: Optional[float] = None,
    preserve_aspect_ratio: bool = True,
    default_dpi: float = 96.0,
) -> Tuple[float, float]:
    """Calculate target width and height in inches given image pixels and optional requested bounds.

    Args:
        px_width: Native pixel width.
        px_height: Native pixel height.
        requested_width: Optional target width in inches.
        requested_height: Optional target height in inches.
        preserve_aspect_ratio: Whether to maintain image aspect ratio.
        default_dpi: Default DPI when dimensions are omitted (default: 96 DPI).

    Returns:
        (final_width_inches, final_height_inches)
    """
    aspect = float(px_width) / float(px_height) if px_height > 0 else 1.0

    if requested_width is None and requested_height is None:
        # Compute from DPI, capping at maximum reasonable defaults (e.g. 8.0 x 6.0)
        calc_w = px_width / default_dpi
        calc_h = px_height / default_dpi
        max_w = 8.0
        max_h = 6.0
        if calc_w > max_w or calc_h > max_h:
            scale = min(max_w / calc_w, max_h / calc_h)
            calc_w *= scale
            calc_h *= scale
        return round(calc_w, 4), round(calc_h, 4)

    if requested_width is not None and requested_height is None:
        calc_w = float(requested_width)
        calc_h = calc_w / aspect
        return round(calc_w, 4), round(calc_h, 4)

    if requested_height is not None and requested_width is None:
        calc_h = float(requested_height)
        calc_w = calc_h * aspect
        return round(calc_w, 4), round(calc_h, 4)

    # Both requested_width and requested_height are provided
    req_w = float(requested_width)
    req_h = float(requested_height)

    if not preserve_aspect_ratio:
        return round(req_w, 4), round(req_h, 4)

    # Fit within bounding box preserving aspect ratio
    scale = min(req_w / px_width, req_h / px_height)
    fit_w = px_width * scale
    fit_h = px_height * scale
    return round(fit_w, 4), round(fit_h, 4)


def add_picture(
    slide: Any,
    image_path: Union[str, Path],
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    preserve_aspect_ratio: bool = True,
) -> Dict[str, Any]:
    """Insert a new picture onto a slide with exact geometry and aspect ratio calculation.

    Args:
        slide: Target pptx Slide object.
        image_path: Absolute or relative path to the image file.
        left: Left coordinate in inches.
        top: Top coordinate in inches.
        width: Optional width in inches.
        height: Optional height in inches.
        preserve_aspect_ratio: Whether to preserve aspect ratio.

    Returns:
        Dictionary containing shape details, final geometry, and image metadata.
    """
    path = Path(image_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    px_w, px_h, img_format, _ = _get_image_metadata(path)
    final_w, final_h = calculate_picture_dimensions(
        px_w, px_h, requested_width=width, requested_height=height, preserve_aspect_ratio=preserve_aspect_ratio
    )

    left_emu = inches_to_emu(left)
    top_emu = inches_to_emu(top)
    width_emu = inches_to_emu(final_w)
    height_emu = inches_to_emu(final_h)

    pic_shape = slide.shapes.add_picture(
        str(path),
        left_emu,
        top_emu,
        width=width_emu,
        height=height_emu,
    )

    return {
        "shape_id": pic_shape.shape_id,
        "name": pic_shape.name,
        "shape_type": "picture",
        "geometry": {
            "x": round(left, 4),
            "y": round(top, 4),
            "width": round(final_w, 4),
            "height": round(final_h, 4),
            "left_emu": left_emu,
            "top_emu": top_emu,
            "width_emu": width_emu,
            "height_emu": height_emu,
        },
        "image_metadata": {
            "path": str(path),
            "format": img_format,
            "pixel_width": px_w,
            "pixel_height": px_h,
            "aspect_ratio": round(float(px_w) / float(px_h), 4) if px_h > 0 else 1.0,
        },
    }


def replace_picture(
    slide: Any,
    shape_id: int,
    image_path: Union[str, Path],
    preserve_geometry: bool = True,
) -> Dict[str, Any]:
    """Replace an existing picture or placeholder shape with a new image.

    If the target shape is already a picture, updates its embedded blip image reference.
    If the target shape is a shape placeholder, inserts the picture into the exact
    same bounds, z-order position, and rotation, and removes the placeholder.

    Args:
        slide: Target pptx Slide object.
        shape_id: Target shape ID to replace.
        image_path: Path to the replacement image.
        preserve_geometry: If True, retains original left, top, width, height, and rotation.

    Returns:
        Dictionary containing shape details and updated image metadata.
    """
    path = Path(image_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Replacement image file not found: {image_path}")

    target_shape = None
    for s in slide.shapes:
        if s.shape_id == shape_id:
            target_shape = s
            break

    if target_shape is None:
        raise ValueError(f"Shape with ID {shape_id} not found on slide")

    px_w, px_h, img_format, _ = _get_image_metadata(path)

    orig_left = int(target_shape.left)
    orig_top = int(target_shape.top)
    orig_width = int(target_shape.width)
    orig_height = int(target_shape.height)
    orig_rotation = getattr(target_shape, "rotation", 0.0) or 0.0

    # Case A: Target shape is already an OpenXML Picture shape
    is_picture = (
        getattr(target_shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE
        or len(target_shape._element.xpath(".//a:blip")) > 0
    )

    if is_picture:
        # Register new image part in slide relations
        image_part, rId = slide.part.get_or_add_image_part(str(path))
        blips = target_shape._element.xpath(".//a:blip")
        if not blips:
            raise RuntimeError(f"Could not locate image blip element for shape ID {shape_id}")

        embed_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        blips[0].set(embed_attr, rId)

        if not preserve_geometry:
            # Recalculate dimensions based on new image aspect ratio
            final_w, final_h = calculate_picture_dimensions(px_w, px_h, preserve_aspect_ratio=True)
            target_shape.width = inches_to_emu(final_w)
            target_shape.height = inches_to_emu(final_h)

        return {
            "success": True,
            "shape_id": target_shape.shape_id,
            "name": target_shape.name,
            "operation": "replace_picture",
            "geometry": {
                "x": emu_to_inches(target_shape.left),
                "y": emu_to_inches(target_shape.top),
                "width": emu_to_inches(target_shape.width),
                "height": emu_to_inches(target_shape.height),
                "rotation": round(getattr(target_shape, "rotation", 0.0) or 0.0, 2),
            },
            "image_metadata": {
                "path": str(path),
                "format": img_format,
                "pixel_width": px_w,
                "pixel_height": px_h,
            },
        }

    # Case B: Target shape is a placeholder shape (e.g. auto shape / rectangle)
    sp_tree = target_shape._element.getparent()
    ref_idx = sp_tree.index(target_shape._element)

    # Determine dimensions for new picture
    if preserve_geometry:
        final_w_emu = orig_width
        final_h_emu = orig_height
    else:
        calc_w, calc_h = calculate_picture_dimensions(
            px_w, px_h, requested_width=emu_to_inches(orig_width), preserve_aspect_ratio=True
        )
        final_w_emu = inches_to_emu(calc_w)
        final_h_emu = inches_to_emu(calc_h)

    new_pic = slide.shapes.add_picture(
        str(path),
        orig_left,
        orig_top,
        width=final_w_emu,
        height=final_h_emu,
    )
    if orig_rotation:
        new_pic.rotation = orig_rotation

    # Place new picture at exact same z-index in shape tree
    new_pic_elem = new_pic._element
    sp_tree.remove(new_pic_elem)
    sp_tree.insert(ref_idx, new_pic_elem)

    # Remove placeholder shape
    sp_tree.remove(target_shape._element)

    return {
        "success": True,
        "shape_id": new_pic.shape_id,
        "replaced_shape_id": shape_id,
        "name": new_pic.name,
        "operation": "replace_placeholder_with_picture",
        "geometry": {
            "x": emu_to_inches(new_pic.left),
            "y": emu_to_inches(new_pic.top),
            "width": emu_to_inches(new_pic.width),
            "height": emu_to_inches(new_pic.height),
            "rotation": round(getattr(new_pic, "rotation", 0.0) or 0.0, 2),
        },
        "image_metadata": {
            "path": str(path),
            "format": img_format,
            "pixel_width": px_w,
            "pixel_height": px_h,
        },
    }