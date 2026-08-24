"""Shape, text, and slide manipulation engine with EMU coordinate precision, run-level style preservation, and relationship-safe element copying."""

import copy
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from powerpoint_mcp.models.shape import (
    AlignmentType,
    BoundingBox,
    EMU_PER_INCH,
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.pptx.ooxml import NAMESPACES

P_NS = f"{{{NAMESPACES['p']}}}"
A_NS = f"{{{NAMESPACES['a']}}}"
R_NS = f"{{{NAMESPACES['r']}}}"

# Map text alignment string / enum to PP_ALIGN
ALIGNMENT_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
    "distribute": PP_ALIGN.DISTRIBUTE,
}


def _resolve_slide(
    slide_or_prs: Any,
    slide_number: Optional[int] = None,
) -> Any:
    """Resolve a python-pptx Slide object from a Slide, Presentation, or file path.

    Args:
        slide_or_prs: python-pptx Slide, Presentation, or path string/Path.
        slide_number: 1-indexed slide number (required if slide_or_prs is Presentation or path).

    Returns:
        The resolved python-pptx Slide object.
    """
    # If already a Slide
    if hasattr(slide_or_prs, "shapes") and not hasattr(slide_or_prs, "slides"):
        return slide_or_prs

    # If Presentation object
    if hasattr(slide_or_prs, "slides"):
        prs = slide_or_prs
    elif isinstance(slide_or_prs, (str, Path)):
        prs = Presentation(str(slide_or_prs))
    else:
        raise TypeError(f"Unsupported slide or presentation type: {type(slide_or_prs).__name__}")

    if slide_number is None or slide_number < 1 or slide_number > len(prs.slides):
        raise ValueError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    return prs.slides[slide_number - 1]


def _resolve_target(
    slide_or_prs: Any,
    arg1: Any,
    arg2: Optional[Any] = None,
) -> Tuple[Any, Any]:
    """Parse flexible positional arguments to resolve (slide, shape).

    Supports:
    - (slide, shape_id) -> returns (slide, shape)
    - (prs, slide_number, shape_id) -> returns (slide, shape)
    - (path, slide_number, shape_id) -> returns (slide, shape)
    """
    # Check if slide_or_prs is a Slide directly
    if hasattr(slide_or_prs, "shapes") and not hasattr(slide_or_prs, "slides"):
        slide = slide_or_prs
        shape_id = int(arg1)
    else:
        # slide_or_prs is a Presentation or path
        slide_number = int(arg1)
        if arg2 is None:
            raise ValueError("shape_id must be provided when passing a Presentation or presentation path")
        shape_id = int(arg2)
        slide = _resolve_slide(slide_or_prs, slide_number)

    # Locate shape on slide by ID
    for shape in slide.shapes:
        if shape.shape_id == shape_id:
            return slide, shape

    raise ValueError(f"Shape with ID {shape_id} not found on slide")


def _hex_to_rgb(hex_str: str) -> Optional[RGBColor]:
    """Parse hex string to RGBColor object."""
    cleaned = hex_str.strip().lstrip("#").upper()
    if len(cleaned) == 3:
        cleaned = "".join(c * 2 for c in cleaned)
    if len(cleaned) == 6:
        try:
            r = int(cleaned[0:2], 16)
            g = int(cleaned[2:4], 16)
            b = int(cleaned[4:6], 16)
            return RGBColor(r, g, b)
        except Exception:
            return None
    return None


def _apply_z_order(slide: Any, shape: Any, z_order: Union[int, str]) -> None:
    """Adjust the visual stacking z-order of a shape element inside `<p:spTree>`."""
    target_elem = shape._element
    sp_tree = target_elem.getparent()
    if sp_tree is None:
        return

    # Find all shape elements in spTree (exclude non-shape properties like nvGrpSpPr, grpSpPr)
    shape_tags = ("sp", "pic", "graphicFrame", "grpSp", "cxnSp")
    shape_elements = [e for e in sp_tree if any(e.tag.endswith(tag) for tag in shape_tags)]

    if not shape_elements or target_elem not in shape_elements:
        return

    # Property header count (insertion baseline for sending to back)
    header_count = 0
    for e in sp_tree:
        if e.tag.endswith(("nvGrpSpPr", "grpSpPr")):
            header_count += 1
        else:
            break

    if isinstance(z_order, str):
        action = z_order.strip().lower()
        if action in ("front", "bring_to_front", "top"):
            sp_tree.remove(target_elem)
            sp_tree.append(target_elem)
        elif action in ("back", "send_to_back", "bottom"):
            sp_tree.remove(target_elem)
            sp_tree.insert(header_count, target_elem)
        elif action in ("forward", "bring_forward"):
            curr_idx = sp_tree.index(target_elem)
            if curr_idx < len(sp_tree) - 1:
                sp_tree.remove(target_elem)
                sp_tree.insert(curr_idx + 1, target_elem)
        elif action in ("backward", "send_backward"):
            curr_idx = sp_tree.index(target_elem)
            if curr_idx > header_count:
                sp_tree.remove(target_elem)
                sp_tree.insert(curr_idx - 1, target_elem)
        else:
            try:
                numeric_z = int(action)
                _apply_z_order(slide, shape, numeric_z)
            except ValueError:
                raise ValueError(f"Unknown z-order action: '{z_order}'")

    elif isinstance(z_order, int):
        sp_tree.remove(target_elem)
        # Update shape elements list after removal
        shape_elements = [e for e in sp_tree if any(e.tag.endswith(tag) for tag in shape_tags)]
        total_shapes = len(shape_elements)

        if total_shapes == 0:
            sp_tree.insert(header_count, target_elem)
        elif z_order <= 0:
            sp_tree.insert(header_count, target_elem)
        elif z_order >= total_shapes:
            sp_tree.append(target_elem)
        else:
            ref_elem = shape_elements[z_order]
            ref_idx = sp_tree.index(ref_elem)
            sp_tree.insert(ref_idx, target_elem)


def modify_shape(
    slide_or_prs: Any,
    arg1: Any,
    arg2: Optional[Any] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    rotation: Optional[float] = None,
    z_order: Optional[Union[int, str]] = None,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
    dwidth: Optional[float] = None,
    dheight: Optional[float] = None,
    drotation: Optional[float] = None,
) -> Dict[str, Any]:
    """Modify shape position, dimensions, rotation, and z-order with exact EMU precision.

    Args:
        slide_or_prs: Slide, Presentation, or file path.
        arg1: shape_id (if slide passed) or slide_number (if prs/path passed).
        arg2: shape_id (if prs/path passed).
        x: Absolute X position in inches.
        y: Absolute Y position in inches.
        width: Absolute width in inches.
        height: Absolute height in inches.
        rotation: Absolute rotation in degrees.
        z_order: Target z-index (int) or action ('bring_to_front', 'send_to_back', 'bring_forward', 'send_backward').
        dx: Relative delta X in inches.
        dy: Relative delta Y in inches.
        dwidth: Relative delta width in inches.
        dheight: Relative delta height in inches.
        drotation: Relative delta rotation in degrees.

    Returns:
        Dictionary containing the updated shape properties.
    """
    slide, shape = _resolve_target(slide_or_prs, arg1, arg2)

    # Apply absolute coordinates / dimensions
    if x is not None:
        shape.left = inches_to_emu(x)
    if y is not None:
        shape.top = inches_to_emu(y)
    if width is not None:
        shape.width = inches_to_emu(width)
    if height is not None:
        shape.height = inches_to_emu(height)

    # Apply relative delta adjustments without floating point drift
    if dx is not None:
        shape.left = int(shape.left) + inches_to_emu(dx)
    if dy is not None:
        shape.top = int(shape.top) + inches_to_emu(dy)
    if dwidth is not None:
        shape.width = int(shape.width) + inches_to_emu(dwidth)
    if dheight is not None:
        shape.height = int(shape.height) + inches_to_emu(dheight)

    # Apply rotation
    if rotation is not None:
        shape.rotation = float(rotation)
    if drotation is not None:
        cur_rot = getattr(shape, "rotation", 0.0) or 0.0
        shape.rotation = (cur_rot + float(drotation)) % 360.0

    # Apply z-order
    if z_order is not None:
        _apply_z_order(slide, shape, z_order)

    return {
        "shape_id": shape.shape_id,
        "name": shape.name,
        "x": emu_to_inches(shape.left),
        "y": emu_to_inches(shape.top),
        "width": emu_to_inches(shape.width),
        "height": emu_to_inches(shape.height),
        "rotation": round(getattr(shape, "rotation", 0.0) or 0.0, 2),
        "left_emu": int(shape.left),
        "top_emu": int(shape.top),
        "width_emu": int(shape.width),
        "height_emu": int(shape.height),
    }


def move_shape(
    slide_or_prs: Any,
    arg1: Any,
    arg2: Optional[Any] = None,
    delta_x_inches: Optional[float] = None,
    delta_y_inches: Optional[float] = None,
    x_inches: Optional[float] = None,
    y_inches: Optional[float] = None,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
) -> Dict[str, Any]:
    """Move a shape by absolute coordinates or relative delta offsets in inches."""
    return modify_shape(
        slide_or_prs,
        arg1,
        arg2,
        x=x_inches if x_inches is not None else x,
        y=y_inches if y_inches is not None else y,
        dx=delta_x_inches if delta_x_inches is not None else dx,
        dy=delta_y_inches if delta_y_inches is not None else dy,
    )


def resize_shape(
    slide_or_prs: Any,
    arg1: Any,
    arg2: Optional[Any] = None,
    width_inches: Optional[float] = None,
    height_inches: Optional[float] = None,
    scale_x: Optional[float] = None,
    scale_y: Optional[float] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    dwidth: Optional[float] = None,
    dheight: Optional[float] = None,
) -> Dict[str, Any]:
    """Resize a shape using absolute dimensions, scaling factors, or relative deltas."""
    slide, shape = _resolve_target(slide_or_prs, arg1, arg2)

    target_w = width_inches if width_inches is not None else width
    target_h = height_inches if height_inches is not None else height

    if scale_x is not None:
        target_w = emu_to_inches(int(round(int(shape.width) * float(scale_x))))
    if scale_y is not None:
        target_h = emu_to_inches(int(round(int(shape.height) * float(scale_y))))

    return modify_shape(
        slide,
        shape.shape_id,
        width=target_w,
        height=target_h,
        dwidth=dwidth,
        dheight=dheight,
    )


def delete_shape(
    slide_or_prs: Any,
    arg1: Any,
    arg2: Optional[Any] = None,
) -> bool:
    """Remove a shape cleanly from its slide's shape tree.

    Args:
        slide_or_prs: Slide, Presentation, or presentation path.
        arg1: shape_id (if slide passed) or slide_number (if prs/path passed).
        arg2: shape_id (if prs/path passed).

    Returns:
        True if the shape was successfully removed.
    """
    slide, shape = _resolve_target(slide_or_prs, arg1, arg2)
    elem = shape._element
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)
        return True
    return False


def modify_text(
    slide_or_prs: Any,
    arg1: Optional[Any] = None,
    arg2: Optional[Any] = None,
    text: Optional[str] = None,
    font_family: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    font_size_pt: Optional[float] = None,
    font_size_delta: Optional[float] = None,
    font_size_scale: Optional[float] = None,
    min_font_size: Optional[float] = None,
    max_font_size: Optional[float] = None,
    min_pt: Optional[float] = None,
    max_pt: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None,
    color_rgb: Optional[str] = None,
    alignment: Optional[Union[str, AlignmentType, PP_ALIGN]] = None,
    paragraph_spacing: Optional[float] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    line_spacing: Optional[float] = None,
    margins: Optional[Dict[str, float]] = None,
    shape_id: Optional[int] = None,
) -> Dict[str, Any]:
    """Modify text content, typography, and paragraph styles with run-level style preservation.

    Supports absolute font sizing, relative deltas, and scaling with min/max bounds.

    Args:
        slide_or_prs: Slide, Presentation, or presentation path.
        arg1: shape_id (if slide passed) or slide_number (if prs/path passed).
        arg2: shape_id (if prs/path passed).
        text: New text content (preserves base typography or applies specified styles).
        font_family / font_name: Target font family name.
        font_size / font_size_pt: Absolute font size in points.
        font_size_delta: Relative point delta (+2, -2) to adjust current font size.
        font_size_scale: Multiplier (e.g. 1.15) to scale current font size.
        min_font_size / min_pt: Lower bound clamp for font size in points.
        max_font_size / max_pt: Upper bound clamp for font size in points.
        bold: Bold weight flag.
        italic: Italic flag.
        underline: Underline flag.
        color / color_rgb: RGB color hex (e.g. '#FF5500' or 'FF5500').
        alignment: Paragraph alignment ('left', 'center', 'right', 'justify').
        paragraph_spacing / space_before: Space before paragraph in points.
        space_after: Space after paragraph in points.
        line_spacing: Line spacing in points.
        margins: Margin dict in inches, e.g. {'left': 0.1, 'right': 0.1, 'top': 0.05, 'bottom': 0.05}.

    Returns:
        Dictionary describing the updated text and typography with original/resulting sizes.
    """
    actual_shape_id = shape_id if shape_id is not None else (arg2 if arg2 is not None else arg1)
    target_arg1 = arg1 if arg1 is not None else actual_shape_id
    slide, shape = _resolve_target(slide_or_prs, target_arg1, actual_shape_id if arg2 is not None else None)

    if not getattr(shape, "has_text_frame", False):
        raise ValueError(f"Shape ID {shape.shape_id} does not have a text frame")

    tf = shape.text_frame

    target_font_name = font_family or font_name
    target_font_size_pt = font_size if font_size is not None else font_size_pt
    eff_min_pt = min_font_size if min_font_size is not None else min_pt
    eff_max_pt = max_font_size if max_font_size is not None else max_pt
    target_color_hex = color or color_rgb
    target_space_before = space_before if space_before is not None else paragraph_spacing

    def _calc_target_font_size(current_pt: Optional[float]) -> Optional[float]:
        if target_font_size_pt is not None:
            res = float(target_font_size_pt)
        elif font_size_scale is not None or font_size_delta is not None:
            base = float(current_pt) if current_pt is not None else 14.0
            res = base
            if font_size_scale is not None:
                res = res * float(font_size_scale)
            if font_size_delta is not None:
                res = res + float(font_size_delta)
        else:
            return current_pt

        if eff_min_pt is not None:
            res = max(res, float(eff_min_pt))
        if eff_max_pt is not None:
            res = min(res, float(eff_max_pt))
        res = max(1.0, round(res, 2))
        return res

    # Resolve PP_ALIGN alignment enum if specified
    pp_align: Optional[PP_ALIGN] = None
    if alignment is not None:
        if isinstance(alignment, PP_ALIGN):
            pp_align = alignment
        elif isinstance(alignment, AlignmentType):
            pp_align = ALIGNMENT_MAP.get(alignment.value.lower())
        elif isinstance(alignment, str):
            pp_align = ALIGNMENT_MAP.get(alignment.strip().lower())

    original_font_sizes: List[float] = []
    resulting_font_sizes: List[float] = []

    # If new text is provided: capture base style and recreate text
    if text is not None:
        # Step 1: Capture base formatting properties from the first populated run
        base_font_name: Optional[str] = None
        base_font_size_pt: Optional[float] = None
        base_bold: Optional[bool] = None
        base_italic: Optional[bool] = None
        base_underline: Optional[bool] = None
        base_color_rgb: Optional[RGBColor] = None
        base_alignment: Optional[PP_ALIGN] = None

        for p in tf.paragraphs:
            if base_alignment is None and p.alignment is not None:
                base_alignment = p.alignment
            for r in p.runs:
                if base_font_name is None and r.font.name:
                    base_font_name = r.font.name
                if base_font_size_pt is None and r.font.size is not None:
                    try:
                        base_font_size_pt = float(r.font.size.pt)
                    except Exception:
                        base_font_size_pt = emu_to_pt(int(r.font.size))
                if base_bold is None and r.font.bold is not None:
                    base_bold = r.font.bold
                if base_italic is None and r.font.italic is not None:
                    base_italic = r.font.italic
                if base_underline is None and r.font.underline is not None:
                    base_underline = r.font.underline
                if base_color_rgb is None:
                    try:
                        if r.font.color and r.font.color.type == 1:
                            base_color_rgb = r.font.color.rgb
                    except Exception:
                        pass

        if base_font_size_pt is not None:
            original_font_sizes.append(base_font_size_pt)

        computed_font_size_pt = _calc_target_font_size(base_font_size_pt)
        if computed_font_size_pt is not None:
            resulting_font_sizes.append(computed_font_size_pt)

        effective_font_name = target_font_name or base_font_name
        effective_font_size = Pt(computed_font_size_pt) if computed_font_size_pt is not None else (
            Pt(base_font_size_pt) if base_font_size_pt is not None else None
        )
        effective_bold = bold if bold is not None else base_bold
        effective_italic = italic if italic is not None else base_italic
        effective_underline = underline if underline is not None else base_underline
        effective_color_rgb = _hex_to_rgb(target_color_hex) if target_color_hex is not None else base_color_rgb
        effective_alignment = pp_align if pp_align is not None else base_alignment

        from copy import deepcopy

        lines = text.split("\n")
        num_existing_paras = len(tf.paragraphs)

        for idx, line in enumerate(lines):
            if idx < num_existing_paras:
                p = tf.paragraphs[idx]
                if p.runs:
                    p.runs[0].text = line
                    for extra_r in p.runs[1:]:
                        r_elem = extra_r._r
                        r_parent = r_elem.getparent()
                        if r_parent is not None:
                            r_parent.remove(r_elem)
                else:
                    p.text = line
            else:
                p = tf.add_paragraph()
                p.text = line
                if idx > 0 and len(tf.paragraphs) > 1:
                    prev_p = tf.paragraphs[idx - 1]
                    p.level = prev_p.level
                    if prev_p._p.pPr is not None and p._p.pPr is None:
                        p._p.insert(0, deepcopy(prev_p._p.pPr))

            # Apply explicit paragraph styles only if explicitly requested
            if pp_align is not None:
                p.alignment = pp_align
            if target_space_before is not None:
                p.space_before = Pt(target_space_before)
            if space_after is not None:
                p.space_after = Pt(space_after)
            if line_spacing is not None:
                p.line_spacing = Pt(line_spacing)

            for r in p.runs:
                if effective_font_name:
                    r.font.name = effective_font_name
                if effective_font_size is not None:
                    r.font.size = effective_font_size
                if effective_bold is not None:
                    r.font.bold = effective_bold
                if effective_italic is not None:
                    r.font.italic = effective_italic
                if effective_underline is not None:
                    r.font.underline = effective_underline
                if effective_color_rgb is not None:
                    r.font.color.rgb = effective_color_rgb

        # Remove any excess paragraphs if new text has fewer lines
        if len(lines) < num_existing_paras:
            for extra_p in list(tf.paragraphs)[len(lines):]:
                p_elem = extra_p._p
                parent = p_elem.getparent()
                if parent is not None:
                    parent.remove(p_elem)

    else:
        # Style-only updates across existing paragraphs and runs
        for p in tf.paragraphs:
            if pp_align is not None:
                p.alignment = pp_align
            if target_space_before is not None:
                p.space_before = Pt(target_space_before)
            if space_after is not None:
                p.space_after = Pt(space_after)
            if line_spacing is not None:
                p.line_spacing = Pt(line_spacing)

            for r in p.runs:
                cur_sz = None
                if r.font.size is not None:
                    try:
                        cur_sz = float(r.font.size.pt)
                    except Exception:
                        cur_sz = emu_to_pt(int(r.font.size))
                elif p.font and p.font.size is not None:
                    try:
                        cur_sz = float(p.font.size.pt)
                    except Exception:
                        cur_sz = emu_to_pt(int(p.font.size))

                if cur_sz is not None:
                    original_font_sizes.append(cur_sz)

                new_sz = _calc_target_font_size(cur_sz)
                if new_sz is not None:
                    resulting_font_sizes.append(new_sz)
                    r.font.size = Pt(new_sz)

                if target_font_name:
                    r.font.name = target_font_name
                if bold is not None:
                    r.font.bold = bold
                if italic is not None:
                    r.font.italic = italic
                if underline is not None:
                    r.font.underline = underline
                if target_color_hex:
                    parsed_color = _hex_to_rgb(target_color_hex)
                    if parsed_color:
                        r.font.color.rgb = parsed_color

    # Update text frame margins if provided
    if margins is not None:
        left_m = margins.get("left", margins.get("left_inches", margins.get("margin_left_inches")))
        right_m = margins.get("right", margins.get("right_inches", margins.get("margin_right_inches")))
        top_m = margins.get("top", margins.get("top_inches", margins.get("margin_top_inches")))
        bottom_m = margins.get("bottom", margins.get("bottom_inches", margins.get("margin_bottom_inches")))

        if left_m is not None:
            tf.margin_left = Inches(float(left_m))
        if right_m is not None:
            tf.margin_right = Inches(float(right_m))
        if top_m is not None:
            tf.margin_top = Inches(float(top_m))
        if bottom_m is not None:
            tf.margin_bottom = Inches(float(bottom_m))

    return {
        "shape_id": shape.shape_id,
        "name": shape.name,
        "text": tf.text,
        "paragraph_count": len(tf.paragraphs),
        "font_size": resulting_font_sizes[0] if resulting_font_sizes else (original_font_sizes[0] if original_font_sizes else None),
        "original_font_size": original_font_sizes[0] if original_font_sizes else None,
        "resulting_font_size": resulting_font_sizes[0] if resulting_font_sizes else None,
        "original_font_sizes": original_font_sizes,
        "resulting_font_sizes": resulting_font_sizes,
        "font_name": target_font_name or (tf.paragraphs[0].runs[0].font.name if tf.paragraphs and tf.paragraphs[0].runs else None),
    }


def scale_slide_typography(
    slide_or_prs: Any,
    slide_number: Optional[int] = None,
    scale_factor: float = 1.0,
    font_size_delta: Optional[float] = None,
    min_pt: Optional[float] = None,
    max_pt: Optional[float] = None,
    include_shape_ids: Optional[List[int]] = None,
    exclude_shape_ids: Optional[List[int]] = None,
    preserve_hierarchy: bool = True,
) -> Dict[str, Any]:
    """Proportionally scale or shift typography across all text-bearing shapes on a slide while preserving hierarchy.

    Args:
        slide_or_prs: Slide or Presentation object or file path.
        slide_number: 1-indexed slide number.
        scale_factor: Multiplier for font sizes (e.g. 1.15 for +15%).
        font_size_delta: Point shift added to font sizes (e.g. +2.0).
        min_pt: Minimum resulting font size clamp in points.
        max_pt: Maximum resulting font size clamp in points.
        include_shape_ids: Optional list of shape IDs to exclusively scale.
        exclude_shape_ids: Optional list of shape IDs to skip.
        preserve_hierarchy: If True, preserves relative sizing differences between runs.

    Returns:
        Summary detailing modified shapes with old/new sizes, skipped shapes, and reasons.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs

    shapes_modified: List[Dict[str, Any]] = []
    shapes_skipped: List[Dict[str, Any]] = []

    inc_set = set(include_shape_ids) if include_shape_ids else None
    exc_set = set(exclude_shape_ids) if exclude_shape_ids else set()

    for shape in slide.shapes:
        sid = shape.shape_id
        if not getattr(shape, "has_text_frame", False):
            continue

        if inc_set is not None and sid not in inc_set:
            shapes_skipped.append({
                "shape_id": sid,
                "name": shape.name,
                "reason": "Not in include_shape_ids",
            })
            continue

        if sid in exc_set:
            shapes_skipped.append({
                "shape_id": sid,
                "name": shape.name,
                "reason": "Explicitly excluded",
            })
            continue

        tf = shape.text_frame
        if not tf.text or not tf.text.strip():
            shapes_skipped.append({
                "shape_id": sid,
                "name": shape.name,
                "reason": "Empty text frame",
            })
            continue

        old_sizes: List[float] = []
        new_sizes: List[float] = []

        for p in tf.paragraphs:
            for r in p.runs:
                cur_sz = None
                if r.font.size is not None:
                    try:
                        cur_sz = float(r.font.size.pt)
                    except Exception:
                        cur_sz = emu_to_pt(int(r.font.size))
                elif p.font and p.font.size is not None:
                    try:
                        cur_sz = float(p.font.size.pt)
                    except Exception:
                        cur_sz = emu_to_pt(int(p.font.size))
                else:
                    cur_sz = 14.0

                old_sizes.append(round(cur_sz, 2))

                # Compute new scaled size
                new_sz = cur_sz * float(scale_factor)
                if font_size_delta is not None:
                    new_sz += float(font_size_delta)

                if min_pt is not None:
                    new_sz = max(new_sz, float(min_pt))
                if max_pt is not None:
                    new_sz = min(new_sz, float(max_pt))
                new_sz = max(1.0, round(new_sz, 2))

                new_sizes.append(new_sz)
                r.font.size = Pt(new_sz)

        shapes_modified.append({
            "shape_id": sid,
            "name": shape.name,
            "text_summary": (tf.text or "").strip()[:80],
            "old_font_sizes": old_sizes,
            "new_font_sizes": new_sizes,
            "old_primary_size": old_sizes[0] if old_sizes else None,
            "new_primary_size": new_sizes[0] if new_sizes else None,
        })

    return {
        "success": True,
        "slide_number": getattr(slide, "slide_number", slide_number or 1),
        "total_shapes_modified": len(shapes_modified),
        "total_shapes_skipped": len(shapes_skipped),
        "scale_factor": scale_factor,
        "font_size_delta": font_size_delta,
        "min_pt": min_pt,
        "max_pt": max_pt,
        "shapes_modified": shapes_modified,
        "shapes_skipped": shapes_skipped,
    }


def _xpath_query(elem: Any, query: str) -> List[Any]:
    """Execute XPath query on python-pptx BaseOxmlElement or raw lxml Element."""
    try:
        return elem.xpath(query)
    except TypeError:
        return elem.xpath(query, namespaces=NAMESPACES)


def copy_shape(
    slide_or_prs: Any,
    source_arg1: Any,
    source_arg2: Optional[Any] = None,
    target_slide_number: Optional[int] = None,
    target_slide: Optional[Any] = None,
    offset_x_inches: float = 0.2,
    offset_y_inches: float = 0.2,
) -> int:
    """Duplicate a shape with XML deep-copy, relationship preservation, and position offset.

    Supports:
    - Same-slide copying: `copy_shape(slide, shape_id, offset_x_inches=0.2, offset_y_inches=0.2)`
    - Cross-slide copying: `copy_shape(prs, source_slide_num, shape_id, target_slide_number=2)`

    Args:
        slide_or_prs: Slide or Presentation.
        source_arg1: shape_id (if slide passed) or source_slide_number (if prs passed).
        source_arg2: shape_id (if prs passed).
        target_slide_number: Target slide number (1-indexed) if copying across slides.
        target_slide: Direct target Slide object (optional).
        offset_x_inches: Horizontal displacement for the cloned shape in inches (default 0.2).
        offset_y_inches: Vertical displacement for the cloned shape in inches (default 0.2).

    Returns:
        The newly assigned integer shape_id of the cloned shape.
    """
    source_slide, source_shape = _resolve_target(slide_or_prs, source_arg1, source_arg2)

    # Determine destination slide
    if target_slide is not None:
        dest_slide = target_slide
    elif target_slide_number is not None:
        dest_slide = _resolve_slide(slide_or_prs, target_slide_number)
    else:
        dest_slide = source_slide

    # 1. Deep-copy the shape's XML element
    new_elem = copy.deepcopy(source_shape._element)

    # 2. Generate a globally unique shape ID on the destination slide
    dest_ids = [s.shape_id for s in dest_slide.shapes] + [s.shape_id for s in source_slide.shapes]
    new_id = (max(dest_ids) if dest_ids else 100) + 1

    # 3. Update `<p:cNvPr id="..." name="...">`
    cnvpr_nodes = _xpath_query(new_elem, ".//p:cNvPr | .//p:cNvGrpSpPr/p:cNvPr")
    if cnvpr_nodes:
        cnvpr = cnvpr_nodes[0]
        cnvpr.set("id", str(new_id))
        old_name = cnvpr.get("name", source_shape.name)
        cnvpr.set("name", f"{old_name} (Copy)")

    # 4. Apply offset to shape coordinates `<a:off x="..." y="..."/>`
    dx_emu = inches_to_emu(offset_x_inches)
    dy_emu = inches_to_emu(offset_y_inches)
    off_nodes = _xpath_query(new_elem, ".//a:xfrm/a:off")
    if off_nodes:
        off = off_nodes[0]
        curr_x = int(off.get("x", "0"))
        curr_y = int(off.get("y", "0"))
        off.set("x", str(curr_x + dx_emu))
        off.set("y", str(curr_y + dy_emu))

    # 5. Relationship duplication (for images, media, hyperlinks)
    if hasattr(source_slide, "part") and hasattr(dest_slide, "part"):
        # Find all relationship attributes in new_elem
        embed_nodes = _xpath_query(new_elem, ".//*[@r:embed]")
        for node in embed_nodes:
            old_rid = node.get(f"{R_NS}embed")
            if old_rid and old_rid in source_slide.part.rels:
                rel = source_slide.part.rels[old_rid]
                if dest_slide.part != source_slide.part:
                    new_rid = dest_slide.part.relate_to(rel.target_part, rel.reltype)
                    node.set(f"{R_NS}embed", new_rid)

        link_nodes = _xpath_query(new_elem, ".//*[@r:link]")
        for node in link_nodes:
            old_rid = node.get(f"{R_NS}link")
            if old_rid and old_rid in source_slide.part.rels:
                rel = source_slide.part.rels[old_rid]
                if dest_slide.part != source_slide.part:
                    new_rid = dest_slide.part.relate_to(rel.target_part, rel.reltype)
                    node.set(f"{R_NS}link", new_rid)

        id_nodes = _xpath_query(new_elem, ".//*[@r:id]")
        for node in id_nodes:
            old_rid = node.get(f"{R_NS}id")
            if old_rid and old_rid in source_slide.part.rels:
                rel = source_slide.part.rels[old_rid]
                if dest_slide.part != source_slide.part:
                    new_rid = dest_slide.part.relate_to(rel.target_part, rel.reltype)
                    node.set(f"{R_NS}id", new_rid)

    # 6. Append new element to destination slide's `<p:spTree>`
    dest_slide.shapes._spTree.append(new_elem)

    return new_id
