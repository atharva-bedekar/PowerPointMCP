"""Deterministic composite diagram generation for PowerPoint MCP (flow diagrams, process chains)."""

from typing import Any, Dict, List, Optional, Union
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches, Pt

from powerpoint_mcp.models.shape import (
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.pptx.editor import _resolve_slide
from powerpoint_mcp.pptx.styles import STYLE_PRESETS, _hex_to_rgb, apply_style_to_shape


def create_flow_diagram(
    slide_or_prs: Any,
    steps: List[Union[str, Dict[str, Any]]],
    slide_number: Optional[int] = None,
    direction: str = "horizontal",
    shape_type: str = "rounded_rectangle",
    start_x: float = 1.0,
    start_y: float = 2.2,
    total_width: Optional[float] = None,
    total_height: Optional[float] = None,
    node_width: Optional[float] = None,
    node_height: Optional[float] = None,
    node_gap: Optional[float] = None,
    style_preset: str = "card_default",
    connector_style: str = "arrow",
    connector_color: str = "#94A3B8",
) -> Dict[str, Any]:
    """Create a high-quality multi-step flow diagram with connecting arrows and typography.

    Args:
        slide_or_prs: Slide or Presentation.
        steps: List of step labels or step dicts ({'title': ..., 'description': ..., 'badge': ...}).
        slide_number: 1-indexed slide number.
        direction: 'horizontal' or 'vertical'.
        shape_type: 'rounded_rectangle', 'rectangle', 'chevron', or 'oval'.
        start_x: Diagram origin X in inches.
        start_y: Diagram origin Y in inches.
        total_width: Total span width in inches (defaults to slide width - 2.0).
        total_height: Total span height in inches.
        node_width: Explicit node width in inches.
        node_height: Explicit node height in inches.
        node_gap: Gap between nodes in inches (defaults to 0.4 in).
        style_preset: Preset style ('card_default', 'card_accent', 'badge_primary', etc.).
        connector_style: 'arrow', 'line', or 'none'.
        connector_color: Hex color for connecting arrows/lines.

    Returns:
        Dictionary containing created node and connector shape IDs and coordinates.
    """
    if not steps or len(steps) < 2:
        raise ValueError("At least 2 steps are required to create a flow diagram")

    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    n = len(steps)

    # Resolve shape type
    shape_type_key = shape_type.strip().lower()
    shape_enum = MSO_SHAPE.ROUNDED_RECTANGLE
    if shape_type_key in ("rectangle", "rect"):
        shape_enum = MSO_SHAPE.RECTANGLE
    elif shape_type_key in ("chevron", "arrow_block"):
        shape_enum = MSO_SHAPE.CHEVRON
    elif shape_type_key in ("circle", "oval"):
        shape_enum = MSO_SHAPE.OVAL

    # Calculate layout geometry
    dir_clean = direction.strip().lower()
    is_horiz = dir_clean in ("horizontal", "h", "x")

    gap_in = node_gap if node_gap is not None else (0.45 if is_horiz else 0.35)
    total_span_w = total_width if total_width is not None else 11.333
    total_span_h = total_height if total_height is not None else 4.0

    if is_horiz:
        calc_node_w = node_width if node_width is not None else max(1.2, (total_span_w - (n - 1) * gap_in) / n)
        calc_node_h = node_height if node_height is not None else 1.8
    else:
        calc_node_w = node_width if node_width is not None else 4.5
        calc_node_h = node_height if node_height is not None else max(0.8, (total_span_h - (n - 1) * gap_in) / n)

    node_shapes = []
    connector_shapes = []

    # Get preset styling
    preset_data = STYLE_PRESETS.get(style_preset.strip().lower(), STYLE_PRESETS["card_default"])

    for i, step_item in enumerate(steps):
        # Step metadata
        if isinstance(step_item, dict):
            step_title = str(step_item.get("title", f"Step {i+1}"))
            step_desc = str(step_item.get("description", step_item.get("subtitle", "")))
            step_badge = str(step_item.get("badge", ""))
        else:
            step_title = str(step_item)
            step_desc = ""
            step_badge = ""

        # Node position
        if is_horiz:
            nx = start_x + i * (calc_node_w + gap_in)
            ny = start_y
        else:
            nx = start_x
            ny = start_y + i * (calc_node_h + gap_in)

        # Create step shape
        shp = slide.shapes.add_shape(
            shape_enum,
            Inches(nx),
            Inches(ny),
            Inches(calc_node_w),
            Inches(calc_node_h),
        )
        shp.name = f"Flow Node {i+1} - {step_title}"

        # Apply preset style
        apply_style_to_shape(
            shp,
            fill_color=preset_data.get("fill_color", "#F8FAFC"),
            line_color=preset_data.get("line_color", "#CBD5E1"),
            line_width_pt=preset_data.get("line_width_pt", 1.0),
        )

        # Format text frame
        tf = shp.text_frame
        tf.word_wrap = True
        p_title = tf.paragraphs[0]
        p_title.text = step_title
        if p_title.runs:
            p_title.runs[0].font.size = Pt(13)
            p_title.runs[0].font.bold = True
            p_title.runs[0].font.color.rgb = _hex_to_rgb(preset_data.get("font_color", "#0F172A"))

        if step_desc:
            p_desc = tf.add_paragraph()
            p_desc.text = step_desc
            p_desc.space_before = Pt(4)
            if p_desc.runs:
                p_desc.runs[0].font.size = Pt(10)
                p_desc.runs[0].font.color.rgb = _hex_to_rgb("#475569")

        node_shapes.append(shp)

        # Add connector between nodes
        if i > 0 and connector_style != "none":
            prev_shp = node_shapes[i - 1]
            if is_horiz:
                # Right arrow between nodes
                arrow_w = gap_in * 0.7
                arrow_h = 0.25
                arrow_x = emu_to_inches(int(prev_shp.left) + int(prev_shp.width)) + (gap_in - arrow_w) / 2
                arrow_y = ny + (calc_node_h - arrow_h) / 2

                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW,
                    Inches(arrow_x),
                    Inches(arrow_y),
                    Inches(arrow_w),
                    Inches(arrow_h),
                )
                arrow.name = f"Flow Arrow {i} to {i+1}"
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = _hex_to_rgb(connector_color)
                arrow.line.fill.background()
                connector_shapes.append(arrow)
            else:
                # Down arrow between nodes
                arrow_w = 0.25
                arrow_h = gap_in * 0.7
                arrow_x = nx + (calc_node_w - arrow_w) / 2
                arrow_y = emu_to_inches(int(prev_shp.top) + int(prev_shp.height)) + (gap_in - arrow_h) / 2

                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.DOWN_ARROW,
                    Inches(arrow_x),
                    Inches(arrow_y),
                    Inches(arrow_w),
                    Inches(arrow_h),
                )
                arrow.name = f"Flow Arrow {i} to {i+1}"
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = _hex_to_rgb(connector_color)
                arrow.line.fill.background()
                connector_shapes.append(arrow)

    return {
        "success": True,
        "step_count": n,
        "direction": direction,
        "node_shape_ids": [s.shape_id for s in node_shapes],
        "connector_shape_ids": [s.shape_id for s in connector_shapes],
        "total_shapes_created": len(node_shapes) + len(connector_shapes),
        "nodes": [
            {
                "shape_id": s.shape_id,
                "name": s.name,
                "x": emu_to_inches(int(s.left)),
                "y": emu_to_inches(int(s.top)),
                "width": emu_to_inches(int(s.width)),
                "height": emu_to_inches(int(s.height)),
            }
            for s in node_shapes
        ],
    }
