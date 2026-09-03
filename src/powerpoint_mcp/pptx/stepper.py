"""Deterministic stepper and breadcrumb component engine for PowerPoint MCP."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple, Union

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from powerpoint_mcp.models.shape import (
    BoundingBox,
    EMU_PER_INCH,
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.pptx.components import detect_slide_components
from powerpoint_mcp.pptx.editor import _delete_shape_from_slide, _find_shape_by_id, _resolve_slide
from powerpoint_mcp.pptx.inspector import inspect_slide
from powerpoint_mcp.pptx.styles import _hex_to_rgb, apply_style_to_shape


# Default clean stepper palette
DEFAULT_ACTIVE_FILL = "#2563EB"       # Vibrant Blue
DEFAULT_ACTIVE_TEXT = "#FFFFFF"       # Crisp White
DEFAULT_INACTIVE_FILL = "#F1F5F9"     # Soft Slate
DEFAULT_INACTIVE_TEXT = "#64748B"     # Slate Gray
DEFAULT_CONNECTOR_COLOR = "#CBD5E1"   # Light Gray Border/Arrow


def create_stepper(
    slide_or_prs: Any,
    steps: List[str],
    slide_number: Optional[int] = None,
    active_step: Optional[str] = None,
    start_x: Optional[float] = None,
    start_y: Optional[float] = None,
    total_width: Optional[float] = None,
    node_height: Optional[float] = None,
    shape_type: str = "rounded_rectangle",
    active_fill: Optional[str] = None,
    inactive_fill: Optional[str] = None,
    active_text_color: Optional[str] = None,
    inactive_text_color: Optional[str] = None,
    connector_color: Optional[str] = None,
    font_family: Optional[str] = None,
    font_size_pt: Optional[float] = None,
    reference_slide_model: Optional[Any] = None,
) -> Dict[str, Any]:
    """Create a high-fidelity stepper/breadcrumb component on a slide.

    Args:
        slide_or_prs: python-pptx Slide or Presentation instance.
        steps: List of step label strings (e.g. ['ANALYZE', 'CONNECT', 'CONFIGURE', 'RUN']).
        slide_number: 1-indexed slide number.
        active_step: Label of the currently active step (case-insensitive match).
        start_x: X coordinate origin in inches (default 0.8 in).
        start_y: Y coordinate origin in inches (default 1.15 in).
        total_width: Total span width in inches (default 11.7 in).
        node_height: Height of each step pill in inches (default 0.42 in).
        shape_type: Shape type for nodes ('rounded_rectangle', 'rectangle', 'chevron').
        active_fill: Fill hex color for active step.
        inactive_fill: Fill hex color for inactive steps.
        active_text_color: Font color hex for active step.
        inactive_text_color: Font color hex for inactive steps.
        connector_color: Hex color for connecting arrows.
        font_family: Font family name.
        font_size_pt: Font size in points.
        reference_slide_model: Optional SlideModel to extract geometry/styles from.

    Returns:
        Structured dictionary containing created shape IDs, step count, and coordinates.
    """
    if not steps or len(steps) < 2:
        raise ValueError("At least 2 steps are required to create a stepper")

    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    n = len(steps)

    # Defaults
    sx = start_x if start_x is not None else 0.80
    sy = start_y if start_y is not None else 1.15
    span_w = total_width if total_width is not None else 11.733
    nh = node_height if node_height is not None else 0.42

    act_fill = active_fill or DEFAULT_ACTIVE_FILL
    inact_fill = inactive_fill or DEFAULT_INACTIVE_FILL
    act_text = active_text_color or DEFAULT_ACTIVE_TEXT
    inact_text = inactive_text_color or DEFAULT_INACTIVE_TEXT
    conn_color = connector_color or DEFAULT_CONNECTOR_COLOR
    f_name = font_family or "Segoe UI"
    f_size = font_size_pt or 10.5

    # If reference slide model provided, inherit stepper styling/geometry if available
    if reference_slide_model:
        ref_comps = detect_slide_components(reference_slide_model, getattr(reference_slide_model, "slide_number", 1))
        for c in ref_comps:
            if c.type_str == "stepper" and c.bbox:
                if start_x is None:
                    sx = c.bbox.left_inches
                if start_y is None:
                    sy = c.bbox.top_inches
                if total_width is None:
                    span_w = c.bbox.width_inches
                if node_height is None:
                    nh = c.bbox.height_inches
                break

    # Determine node shape enum
    shape_enum = MSO_SHAPE.ROUNDED_RECTANGLE
    st_clean = shape_type.strip().lower()
    if st_clean in ("rectangle", "rect"):
        shape_enum = MSO_SHAPE.RECTANGLE
    elif st_clean in ("chevron", "arrow"):
        shape_enum = MSO_SHAPE.CHEVRON

    # Compute step widths and gaps
    gap_in = 0.28
    node_w = max(1.0, (span_w - (n - 1) * gap_in) / n)

    # Active step matching
    active_clean = (active_step or steps[0]).strip().upper()

    node_shapes = []
    connector_shapes = []

    for i, step_label in enumerate(steps):
        lbl_clean = step_label.strip()
        is_active = (lbl_clean.upper() == active_clean) or (active_step is None and i == 0)

        cur_x = sx + i * (node_w + gap_in)
        cur_y = sy

        # Create step pill
        shp = slide.shapes.add_shape(
            shape_enum,
            Inches(cur_x),
            Inches(cur_y),
            Inches(node_w),
            Inches(nh),
        )
        shp.name = f"Stepper Step {i+1} - {lbl_clean}" + (" (Active)" if is_active else "")

        # Apply fills and borders
        shp.fill.solid()
        if is_active:
            shp.fill.fore_color.rgb = _hex_to_rgb(act_fill)
            shp.line.color.rgb = _hex_to_rgb(act_fill)
            shp.line.width = Pt(1.0)
        else:
            shp.fill.fore_color.rgb = _hex_to_rgb(inact_fill)
            shp.line.color.rgb = _hex_to_rgb(DEFAULT_CONNECTOR_COLOR)
            shp.line.width = Pt(1.0)

        # Text formatting
        tf = shp.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)

        p = tf.paragraphs[0]
        p.text = lbl_clean
        p.alignment = 2  # Center (PP_ALIGN.CENTER)
        if p.runs:
            r = p.runs[0]
            r.font.name = f_name
            r.font.size = Pt(f_size)
            r.font.bold = is_active
            r.font.color.rgb = _hex_to_rgb(act_text if is_active else inact_text)

        node_shapes.append(shp)

        # Add connector arrow between steps
        if i > 0:
            arrow_w = gap_in * 0.55
            arrow_h = 0.16
            arrow_x = (cur_x - gap_in) + (gap_in - arrow_w) / 2
            arrow_y = cur_y + (nh - arrow_h) / 2

            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(arrow_x),
                Inches(arrow_y),
                Inches(arrow_w),
                Inches(arrow_h),
            )
            arrow.name = f"Stepper Connector {i} to {i+1}"
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = _hex_to_rgb(conn_color)
            arrow.line.fill.background()
            connector_shapes.append(arrow)

    all_created_ids = [s.shape_id for s in node_shapes] + [s.shape_id for s in connector_shapes]

    return {
        "success": True,
        "step_count": n,
        "steps": steps,
        "active_step": active_step or steps[0],
        "node_shape_ids": [s.shape_id for s in node_shapes],
        "connector_shape_ids": [s.shape_id for s in connector_shapes],
        "all_shape_ids": all_created_ids,
        "bbox": {
            "left": round(sx, 4),
            "top": round(sy, 4),
            "width": round(span_w, 4),
            "height": round(nh, 4),
        },
    }


def update_stepper(
    slide_or_prs: Any,
    active_step: str,
    slide_number: Optional[int] = None,
    steps: Optional[List[str]] = None,
    reference_slide_model: Optional[Any] = None,
    active_fill: Optional[str] = None,
    inactive_fill: Optional[str] = None,
) -> Dict[str, Any]:
    """Update active step on an existing slide stepper component.

    Completely replaces old stepper shapes (nodes, text, and connectors) cleanly
    to ensure zero orphaned shapes or duplicate geometry remain.

    Args:
        slide_or_prs: python-pptx Slide or Presentation instance.
        active_step: Label of the new active step.
        slide_number: 1-indexed slide number.
        steps: Optional override list of steps.
        reference_slide_model: Optional reference SlideModel for styling.
        active_fill: Optional active fill color hex.
        inactive_fill: Optional inactive fill color hex.

    Returns:
        Structured dictionary confirming updated active step and shape IDs.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    s_num = slide_number or 1

    # Detect existing stepper on slide
    existing_comps = detect_slide_components(slide_or_prs, s_num)
    stepper_comp = None
    for c in existing_comps:
        if c.type_str == "stepper":
            stepper_comp = c
            break

    existing_steps: List[str] = []
    sx: Optional[float] = None
    sy: Optional[float] = None
    span_w: Optional[float] = None
    nh: Optional[float] = None

    if stepper_comp:
        existing_steps = stepper_comp.properties.get("steps", [])
        if stepper_comp.bbox:
            sx = stepper_comp.bbox.left_inches
            sy = stepper_comp.bbox.top_inches
            span_w = stepper_comp.bbox.width_inches
            nh = stepper_comp.bbox.height_inches

        # Delete all old stepper constituent shapes cleanly
        for sid in stepper_comp.shape_ids:
            try:
                _delete_shape_from_slide(slide, sid)
            except Exception:
                pass

    final_steps = steps if (steps and len(steps) >= 2) else (existing_steps if len(existing_steps) >= 2 else ["STEP 1", "STEP 2"])

    # Create new clean stepper in the exact same geometry with new active step
    return create_stepper(
        slide_or_prs=slide,
        steps=final_steps,
        slide_number=s_num,
        active_step=active_step,
        start_x=sx,
        start_y=sy,
        total_width=span_w,
        node_height=nh,
        active_fill=active_fill,
        inactive_fill=inactive_fill,
        reference_slide_model=reference_slide_model,
    )
