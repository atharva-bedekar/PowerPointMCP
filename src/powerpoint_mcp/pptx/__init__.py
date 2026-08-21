"""PPTX inspection, styles, and relationships extraction utilities."""

from powerpoint_mcp.pptx.inspector import (
    PPTXInspector,
    infer_semantic_role,
    inspect_presentation,
    inspect_shape,
    inspect_slide,
    map_shape_type,
    match_shapes,
)
from powerpoint_mcp.pptx.relationships import (
    extract_embedded_images,
    extract_hyperlinks,
    get_image_part_from_shape,
    inspect_slide_relationships,
)
from powerpoint_mcp.pptx.styles import (
    extract_alignment_name,
    extract_fill_style,
    extract_font_style,
    extract_line_style,
    extract_paragraph,
    extract_rgb_hex,
    extract_run,
    extract_shape_properties,
    extract_text_frame,
    extract_vertical_anchor_name,
)

__all__ = [
    "PPTXInspector",
    "extract_alignment_name",
    "extract_embedded_images",
    "extract_fill_style",
    "extract_font_style",
    "extract_hyperlinks",
    "extract_line_style",
    "extract_paragraph",
    "extract_rgb_hex",
    "extract_run",
    "extract_shape_properties",
    "extract_text_frame",
    "extract_vertical_anchor_name",
    "get_image_part_from_shape",
    "infer_semantic_role",
    "inspect_presentation",
    "inspect_shape",
    "inspect_slide",
    "inspect_slide_relationships",
    "map_shape_type",
    "match_shapes",
]
