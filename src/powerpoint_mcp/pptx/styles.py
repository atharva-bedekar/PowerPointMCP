"""Helpers for extracting, serializing, and formatting PowerPoint styles, fills, lines, and typography."""

from typing import Any, Dict, List, Optional, Union
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from powerpoint_mcp.models.shape import (
    ParagraphModel,
    TextFrameModel,
    TextRunModel,
    TextStyle,
    emu_to_inches,
    emu_to_pt,
)


def extract_rgb_hex(color_format: Any) -> Optional[str]:
    """Safely extract hex RGB color string ("RRGGBB" or "#RRGGBB") from a ColorFormat object."""
    if color_format is None:
        return None
    try:
        color_type = getattr(color_format, "type", None)
        if color_type == MSO_COLOR_TYPE.RGB or color_type == 1:
            rgb = color_format.rgb
            if rgb is not None:
                return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        elif color_type == MSO_COLOR_TYPE.THEME or color_type == 2:
            theme_color = getattr(color_format, "theme_color", None)
            if theme_color is not None:
                return f"theme:{theme_color.name if hasattr(theme_color, 'name') else str(theme_color)}"
        # Fallback inspection of raw XML element if RGB is accessible directly
        try:
            rgb = color_format.rgb
            if rgb is not None:
                return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        except Exception:
            pass
    except Exception:
        pass
    return None


def extract_font_style(font: Any, fallback_font: Any = None) -> TextStyle:
    """Extract TextStyle model from a python-pptx Font object with optional fallback font."""
    font_name: Optional[str] = getattr(font, "name", None) if font else None
    if not font_name and fallback_font:
        font_name = getattr(fallback_font, "name", None)

    font_size_pt: Optional[float] = None
    if font and font.size is not None:
        try:
            font_size_pt = round(float(font.size.pt), 2)
        except Exception:
            font_size_pt = emu_to_pt(int(font.size))
    elif fallback_font and fallback_font.size is not None:
        try:
            font_size_pt = round(float(fallback_font.size.pt), 2)
        except Exception:
            font_size_pt = emu_to_pt(int(fallback_font.size))

    bold: Optional[bool] = getattr(font, "bold", None) if font else None
    if bold is None and fallback_font:
        bold = getattr(fallback_font, "bold", None)

    italic: Optional[bool] = getattr(font, "italic", None) if font else None
    if italic is None and fallback_font:
        italic = getattr(fallback_font, "italic", None)

    underline: Optional[bool] = getattr(font, "underline", None) if font else None
    if underline is None and fallback_font:
        underline = getattr(fallback_font, "underline", None)

    color_rgb: Optional[str] = None
    try:
        if font and hasattr(font, "color"):
            color_rgb = extract_rgb_hex(font.color)
        if color_rgb is None and fallback_font and hasattr(fallback_font, "color"):
            color_rgb = extract_rgb_hex(fallback_font.color)
    except Exception:
        pass

    return TextStyle(
        font_name=font_name,
        font_size_pt=font_size_pt,
        bold=bold,
        italic=italic,
        underline=underline,
        color_rgb=color_rgb,
    )


def extract_alignment_name(alignment_val: Any) -> Optional[str]:
    """Map python-pptx PP_ALIGN enum to standard string."""
    if alignment_val is None:
        return None
    mapping = {
        PP_ALIGN.LEFT: "left",
        PP_ALIGN.CENTER: "center",
        PP_ALIGN.RIGHT: "right",
        PP_ALIGN.JUSTIFY: "justify",
        PP_ALIGN.DISTRIBUTE: "distribute",
    }
    return mapping.get(alignment_val, str(alignment_val).lower())


def extract_vertical_anchor_name(anchor_val: Any) -> Optional[str]:
    """Map python-pptx MSO_ANCHOR enum to standard string."""
    if anchor_val is None:
        return None
    mapping = {
        MSO_ANCHOR.TOP: "top",
        MSO_ANCHOR.MIDDLE: "middle",
        MSO_ANCHOR.BOTTOM: "bottom",
    }
    return mapping.get(anchor_val, str(anchor_val).lower())


def extract_run(run: Any, fallback_font: Any = None) -> TextRunModel:
    """Extract TextRunModel from a python-pptx Run object."""
    r_font = getattr(run, "font", None)
    style = extract_font_style(r_font, fallback_font=fallback_font)
    hyperlink_target: Optional[str] = None
    try:
        if hasattr(run, "hyperlink") and run.hyperlink is not None:
            hyperlink_target = getattr(run.hyperlink, "address", None)
    except Exception:
        pass

    return TextRunModel(
        text=run.text or "",
        style=style,
        hyperlink_target=hyperlink_target,
    )


def extract_paragraph(paragraph: Any) -> ParagraphModel:
    """Extract ParagraphModel from a python-pptx Paragraph object."""
    p_font = getattr(paragraph, "font", None)
    runs: List[TextRunModel] = []
    for run in paragraph.runs:
        runs.append(extract_run(run, fallback_font=p_font))

    # If no runs exist but paragraph has text
    if not runs and paragraph.text:
        runs.append(TextRunModel(text=paragraph.text, style=extract_font_style(p_font)))

    alignment = extract_alignment_name(getattr(paragraph, "alignment", None))
    level = getattr(paragraph, "level", 0) or 0

    line_spacing_pt: Optional[float] = None
    if paragraph.line_spacing is not None:
        try:
            if hasattr(paragraph.line_spacing, "pt"):
                line_spacing_pt = round(float(paragraph.line_spacing.pt), 2)
            else:
                line_spacing_pt = round(float(paragraph.line_spacing) * 12.0, 2)
        except Exception:
            pass

    space_before_pt: Optional[float] = None
    if paragraph.space_before is not None:
        try:
            space_before_pt = round(float(paragraph.space_before.pt), 2)
        except Exception:
            pass

    space_after_pt: Optional[float] = None
    if paragraph.space_after is not None:
        try:
            space_after_pt = round(float(paragraph.space_after.pt), 2)
        except Exception:
            pass

    return ParagraphModel(
        text=paragraph.text or "",
        runs=runs,
        alignment=alignment,
        level=level,
        line_spacing_pt=line_spacing_pt,
        space_before_pt=space_before_pt,
        space_after_pt=space_after_pt,
    )


def extract_text_frame(text_frame: Any) -> TextFrameModel:
    """Extract TextFrameModel from a python-pptx TextFrame object."""
    if text_frame is None:
        return TextFrameModel(text="")

    paragraphs: List[ParagraphModel] = []
    for p in text_frame.paragraphs:
        paragraphs.append(extract_paragraph(p))

    word_wrap = getattr(text_frame, "word_wrap", True)
    if word_wrap is None:
        word_wrap = True

    margin_left = 0.1
    margin_right = 0.1
    margin_top = 0.05
    margin_bottom = 0.05

    try:
        if text_frame.margin_left is not None:
            margin_left = emu_to_inches(int(text_frame.margin_left))
    except Exception:
        pass

    try:
        if text_frame.margin_right is not None:
            margin_right = emu_to_inches(int(text_frame.margin_right))
    except Exception:
        pass

    try:
        if text_frame.margin_top is not None:
            margin_top = emu_to_inches(int(text_frame.margin_top))
    except Exception:
        pass

    try:
        if text_frame.margin_bottom is not None:
            margin_bottom = emu_to_inches(int(text_frame.margin_bottom))
    except Exception:
        pass

    vertical_anchor = extract_vertical_anchor_name(getattr(text_frame, "vertical_anchor", None))

    return TextFrameModel(
        text=text_frame.text or "",
        paragraphs=paragraphs,
        word_wrap=word_wrap,
        margin_left_inches=margin_left,
        margin_right_inches=margin_right,
        margin_top_inches=margin_top,
        margin_bottom_inches=margin_bottom,
        vertical_anchor=vertical_anchor,
    )


def extract_fill_style(shape: Any) -> Dict[str, Any]:
    """Safely extract fill formatting details from a python-pptx Shape or FillFormat."""
    res: Dict[str, Any] = {
        "type": "none",
        "color": None,
    }
    if not hasattr(shape, "fill"):
        return res

    try:
        fill = shape.fill
        fill_type = getattr(fill, "type", None)
        if fill_type == MSO_FILL.SOLID or fill_type == 1:
            res["type"] = "solid"
            res["color"] = extract_rgb_hex(getattr(fill, "fore_color", None))
        elif fill_type == MSO_FILL.GRADIENT or fill_type == 3:
            res["type"] = "gradient"
        elif fill_type == MSO_FILL.PATTERNED or fill_type == 2:
            res["type"] = "pattern"
        elif fill_type == MSO_FILL.PICTURE or fill_type == 6:
            res["type"] = "picture"
        elif fill_type == MSO_FILL.BACKGROUND or fill_type == 5:
            res["type"] = "background"
        elif fill_type is None:
            res["type"] = "none"
        else:
            res["type"] = "other"
    except Exception:
        pass

    return res


def extract_line_style(shape: Any) -> Dict[str, Any]:
    """Safely extract line border formatting details from a python-pptx Shape or LineFormat."""
    res: Dict[str, Any] = {
        "color": None,
        "width_pt": None,
        "fill_type": "none",
    }
    if not hasattr(shape, "line"):
        return res

    try:
        line = shape.line
        fill_type = getattr(line, "fill", None)
        if fill_type is not None and hasattr(fill_type, "type"):
            if fill_type.type == MSO_FILL.SOLID or fill_type.type == 1:
                res["fill_type"] = "solid"
                res["color"] = extract_rgb_hex(getattr(fill_type, "fore_color", None))
        else:
            # Try line.color directly
            try:
                if hasattr(line, "color"):
                    res["color"] = extract_rgb_hex(line.color)
                    if res["color"]:
                        res["fill_type"] = "solid"
            except Exception:
                pass

        if getattr(line, "width", None) is not None:
            try:
                res["width_pt"] = round(float(line.width.pt), 2)
            except Exception:
                res["width_pt"] = emu_to_pt(int(line.width))
    except Exception:
        pass

    return res


def extract_shape_properties(shape: Any) -> Dict[str, Any]:
    """Extract extended properties of a shape including placeholder details, tables, charts, or media."""
    props: Dict[str, Any] = {}
    try:
        if shape.is_placeholder:
            props["is_placeholder"] = True
            ph_format = shape.placeholder_format
            props["placeholder_type"] = ph_format.type.name if hasattr(ph_format.type, "name") else str(ph_format.type)
            props["placeholder_idx"] = getattr(ph_format, "idx", None)
    except Exception:
        pass

    if getattr(shape, "has_table", False):
        try:
            table = shape.table
            props["table"] = {
                "rows": len(table.rows),
                "columns": len(table.columns),
            }
        except Exception:
            pass

    if getattr(shape, "has_chart", False):
        try:
            chart = shape.chart
            props["chart"] = {
                "chart_title": chart.chart_title.text_frame.text if chart.has_title else None,
                "chart_type": str(chart.chart_type),
            }
        except Exception:
            pass

    return props
