"""Structured Card and Card List component generation for PowerPoint MCP."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

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
from powerpoint_mcp.pptx.editor import _resolve_slide
from powerpoint_mcp.pptx.inspector import inspect_slide
from powerpoint_mcp.pptx.styles import STYLE_PRESETS, _hex_to_rgb, apply_style_to_shape


def create_structured_card_list(
    slide_or_prs: Any,
    container_bbox: Dict[str, float],
    items: List[Dict[str, Any]],
    slide_number: Optional[int] = None,
    divider: bool = True,
    reference_slide_model: Optional[Any] = None,
    style_preset: str = "card_default",
    container_fill: Optional[str] = None,
    container_border: Optional[str] = None,
    title_color: Optional[str] = None,
    desc_color: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a structured container card with organized item rows and optional horizontal dividers.

    Args:
        slide_or_prs: Slide or Presentation instance.
        container_bbox: Dict with left, top, width, height in inches.
        items: List of item dicts with 'title' and optional 'description', 'subtitle', 'badge'.
        slide_number: 1-indexed slide number.
        divider: Whether to insert subtle dividing lines between item rows.
        reference_slide_model: Optional reference SlideModel to inherit typography and styling.
        style_preset: Preset style name ('card_default', 'card_accent', etc.).
        container_fill: Override background fill hex color.
        container_border: Override border stroke hex color.
        title_color: Override item title font color hex.
        desc_color: Override item description font color hex.

    Returns:
        Structured dictionary containing created container, row textboxes, and divider shape IDs.
    """
    if not items:
        raise ValueError("items list must contain at least one item")

    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs

    left_in = float(container_bbox.get("left", 1.0))
    top_in = float(container_bbox.get("top", 1.8))
    width_in = float(container_bbox.get("width", 8.0))
    height_in = float(container_bbox.get("height", 4.0))

    preset_data = STYLE_PRESETS.get(style_preset.strip().lower(), STYLE_PRESETS["card_default"])

    c_fill = container_fill or preset_data.get("fill_color", "#FFFFFF")
    c_border = container_border or preset_data.get("line_color", "#E2E8F0")
    t_color = title_color or preset_data.get("font_color", "#0F172A")
    d_color = desc_color or "#475569"

    # 1. Create main container shape
    container_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )
    container_shape.name = "Structured Card Container"
    apply_style_to_shape(
        container_shape,
        fill_color=c_fill,
        line_color=c_border,
        line_width_pt=1.0,
    )

    n_items = len(items)
    padding_x = 0.30
    padding_y = 0.25
    inner_w = max(1.0, width_in - 2 * padding_x)
    inner_h = max(1.0, height_in - 2 * padding_y)
    row_h = inner_h / n_items

    created_shapes: List[Any] = [container_shape]
    divider_shapes: List[Any] = []
    item_shapes: List[Any] = []

    # 2. Add each item row
    for i, itm in enumerate(items):
        row_top = top_in + padding_y + i * row_h
        itm_title = str(itm.get("title", f"Item {i+1}"))
        itm_desc = str(itm.get("description", itm.get("body", "")))

        # Create text box for item
        tb = slide.shapes.add_textbox(
            Inches(left_in + padding_x),
            Inches(row_top),
            Inches(inner_w),
            Inches(row_h * 0.90),
        )
        tb.name = f"Card Item {i+1} - {itm_title}"
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.0)
        tf.margin_top = Inches(0.0)
        tf.margin_right = Inches(0.0)
        tf.margin_bottom = Inches(0.0)

        # Title paragraph
        p_t = tf.paragraphs[0]
        p_t.text = itm_title
        if p_t.runs:
            p_t.runs[0].font.size = Pt(13)
            p_t.runs[0].font.bold = True
            p_t.runs[0].font.name = "Segoe UI"
            p_t.runs[0].font.color.rgb = _hex_to_rgb(t_color)

        # Description paragraph
        if itm_desc:
            p_d = tf.add_paragraph()
            p_d.text = itm_desc
            p_d.space_before = Pt(3)
            if p_d.runs:
                p_d.runs[0].font.size = Pt(10.5)
                p_d.runs[0].font.name = "Segoe UI"
                p_d.runs[0].font.color.rgb = _hex_to_rgb(d_color)

        item_shapes.append(tb)
        created_shapes.append(tb)

        # 3. Add horizontal divider between rows
        if divider and i < n_items - 1:
            div_y = row_top + row_h
            div = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(left_in + padding_x),
                Inches(div_y),
                Inches(inner_w),
                Inches(0.01),  # Hairline
            )
            div.name = f"Card Divider {i+1}"
            div.fill.solid()
            div.fill.fore_color.rgb = _hex_to_rgb("#F1F5F9")
            div.line.fill.background()
            divider_shapes.append(div)
            created_shapes.append(div)

    return {
        "success": True,
        "container_shape_id": container_shape.shape_id,
        "item_shape_ids": [s.shape_id for s in item_shapes],
        "divider_shape_ids": [s.shape_id for s in divider_shapes],
        "all_shape_ids": [s.shape_id for s in created_shapes],
        "item_count": n_items,
        "bbox": {
            "left": round(left_in, 4),
            "top": round(top_in, 4),
            "width": round(width_in, 4),
            "height": round(height_in, 4),
        },
    }
