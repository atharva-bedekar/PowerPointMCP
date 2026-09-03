"""Component-level atomic geometric operations (move, resize) for PowerPoint MCP."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

from pptx import Presentation
from pptx.util import Inches

from powerpoint_mcp.models.shape import (
    BoundingBox,
    EMU_PER_INCH,
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.pptx.components import detect_slide_components
from powerpoint_mcp.pptx.editor import _find_shape_by_id, _resolve_slide
from powerpoint_mcp.pptx.inspector import inspect_slide


def move_component(
    slide_or_prs: Any,
    component_id: str,
    slide_number: Optional[int] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
) -> Dict[str, Any]:
    """Move all constituent shapes of a component atomically preserving internal relative offsets.

    Args:
        slide_or_prs: Slide or Presentation instance.
        component_id: Target component identifier or type ('header', 'stepper', 'card_1', 'content_area', etc.).
        slide_number: 1-indexed slide number.
        x: Absolute destination X coordinate in inches for component top-left.
        y: Absolute destination Y coordinate in inches for component top-left.
        dx: Relative delta shift X in inches.
        dy: Relative delta shift Y in inches.

    Returns:
        Structured dictionary confirming modified component ID, shape IDs, and new bounding box.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    s_num = slide_number or 1

    comps = detect_slide_components(slide_or_prs, s_num)
    target_comp = None
    for c in comps:
        if c.component_id.lower() == component_id.lower() or c.type_str == component_id.lower() or c.component_id.lower().startswith(component_id.lower()):
            target_comp = c
            break

    if not target_comp and component_id.lower() in ("content_area", "content_container", "layout", "cards", "card"):
        for c in comps:
            if c.type_str in ("card", "card_list", "content_area", "content_container"):
                target_comp = c
                break

    if not target_comp:
        raise ValueError(f"Component '{component_id}' not found on slide {s_num}")

    if not target_comp.bbox:
        raise ValueError(f"Component '{component_id}' has no valid bounding box")

    old_bbox = target_comp.bbox

    # Calculate shift deltas
    shift_dx = 0.0
    shift_dy = 0.0

    if dx is not None:
        shift_dx = dx
    elif x is not None:
        shift_dx = x - old_bbox.left_inches

    if dy is not None:
        shift_dy = dy
    elif y is not None:
        shift_dy = y - old_bbox.top_inches

    delta_emu_x = inches_to_emu(shift_dx)
    delta_emu_y = inches_to_emu(shift_dy)

    modified_shapes = []
    for sid in target_comp.shape_ids:
        pt_shape = _find_shape_by_id(slide, sid)
        if pt_shape:
            pt_shape.left = int(pt_shape.left) + delta_emu_x
            pt_shape.top = int(pt_shape.top) + delta_emu_y
            modified_shapes.append(sid)

    new_left = old_bbox.left_inches + shift_dx
    new_top = old_bbox.top_inches + shift_dy

    return {
        "success": True,
        "component_id": target_comp.component_id,
        "component_type": target_comp.type_str,
        "slide_number": s_num,
        "shape_ids": modified_shapes,
        "shift_dx": round(shift_dx, 4),
        "shift_dy": round(shift_dy, 4),
        "new_bbox": {
            "left": round(new_left, 4),
            "top": round(new_top, 4),
            "width": round(old_bbox.width_inches, 4),
            "height": round(old_bbox.height_inches, 4),
        },
    }


def resize_component(
    slide_or_prs: Any,
    component_id: str,
    slide_number: Optional[int] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    dwidth: Optional[float] = None,
    dheight: Optional[float] = None,
    scale_width: Optional[float] = None,
    scale_height: Optional[float] = None,
    reflow_children: bool = True,
) -> Dict[str, Any]:
    """Resize a component and proportionally adjust/reflow all its constituent shapes atomically.

    Args:
        slide_or_prs: Slide or Presentation instance.
        component_id: Target component identifier or type ('content_area', 'card_1', 'card_list', etc.).
        slide_number: 1-indexed slide number.
        width: Absolute target width in inches.
        height: Absolute target height in inches.
        dwidth: Relative delta width in inches.
        dheight: Relative delta height in inches.
        scale_width: Width scale multiplier (e.g. 1.15).
        scale_height: Height scale multiplier (e.g. 1.10).
        reflow_children: Whether to proportionally adjust child shape positions and sizes.

    Returns:
        Structured dictionary confirming modified shape IDs and updated bounding box.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    s_num = slide_number or 1

    comps = detect_slide_components(slide_or_prs, s_num)
    target_comp = None
    for c in comps:
        if c.component_id.lower() == component_id.lower() or c.type_str == component_id.lower() or c.component_id.lower().startswith(component_id.lower()):
            target_comp = c
            break

    if not target_comp and component_id.lower() in ("content_area", "content_container", "layout", "cards", "card"):
        for c in comps:
            if c.type_str in ("card", "card_list", "content_area", "content_container"):
                target_comp = c
                break

    if not target_comp:
        raise ValueError(f"Component '{component_id}' not found on slide {s_num}")

    if not target_comp.bbox:
        raise ValueError(f"Component '{component_id}' has no valid bounding box")

    old_b = target_comp.bbox
    old_w = old_b.width_inches
    old_h = old_b.height_inches

    # Determine target dimensions
    if width is not None:
        new_w = max(0.2, width)
    elif scale_width is not None:
        new_w = max(0.2, old_w * scale_width)
    elif dwidth is not None:
        new_w = max(0.2, old_w + dwidth)
    else:
        new_w = old_w

    if height is not None:
        new_h = max(0.2, height)
    elif scale_height is not None:
        new_h = max(0.2, old_h * scale_height)
    elif dheight is not None:
        new_h = max(0.2, old_h + dheight)
    else:
        new_h = old_h

    factor_x = new_w / old_w if old_w > 0 else 1.0
    factor_y = new_h / old_h if old_h > 0 else 1.0

    origin_x = old_b.left_inches
    origin_y = old_b.top_inches

    modified_shapes = []
    for sid in target_comp.shape_ids:
        pt_shape = _find_shape_by_id(slide, sid)
        if pt_shape:
            cur_l = emu_to_inches(int(pt_shape.left))
            cur_t = emu_to_inches(int(pt_shape.top))
            cur_w = emu_to_inches(int(pt_shape.width))
            cur_h = emu_to_inches(int(pt_shape.height))

            if reflow_children:
                # Scale relative offset from top-left origin
                rel_x = cur_l - origin_x
                rel_y = cur_t - origin_y

                new_rel_x = rel_x * factor_x
                new_rel_y = rel_y * factor_y

                pt_shape.left = Inches(origin_x + new_rel_x)
                pt_shape.top = Inches(origin_y + new_rel_y)
                pt_shape.width = Inches(max(0.1, cur_w * factor_x))
                pt_shape.height = Inches(max(0.1, cur_h * factor_y))
            else:
                pt_shape.width = Inches(max(0.1, cur_w * factor_x))
                pt_shape.height = Inches(max(0.1, cur_h * factor_y))

            modified_shapes.append(sid)

    return {
        "success": True,
        "component_id": target_comp.component_id,
        "component_type": target_comp.type_str,
        "slide_number": s_num,
        "shape_ids": modified_shapes,
        "new_bbox": {
            "left": round(origin_x, 4),
            "top": round(origin_y, 4),
            "width": round(new_w, 4),
            "height": round(new_h, 4),
        },
    }
