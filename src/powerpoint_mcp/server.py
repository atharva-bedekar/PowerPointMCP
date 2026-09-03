"""PowerPoint MCP Server providing 19 tools and 3 resources for Antigravity."""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys
from typing import Any, Dict, List, Optional, Union

import anyio
from mcp.server.mcpserver import MCPServer

from powerpoint_mcp.tools.editing import (
    ppt_add_picture,
    ppt_align_shapes,
    ppt_apply_style,
    ppt_batch_modify_shapes,
    ppt_batch_modify_table_cells,
    ppt_batch_modify_tables,
    ppt_batch_modify_text,
    ppt_copy_shape,
    ppt_create_flow_diagram,
    ppt_create_stepper,
    ppt_create_structured_card_list,
    ppt_delete_shape,
    ppt_distribute_shapes,
    ppt_equalize_sizes,
    ppt_merge_table_cells,
    ppt_modify_ooxml,
    ppt_modify_shape,
    ppt_modify_text,
    ppt_move_component,
    ppt_move_container,
    ppt_move_shape,
    ppt_reflow_container,
    ppt_replace_picture,
    ppt_resize_component,
    ppt_resize_container,
    ppt_resize_shape,
    ppt_scale_slide_typography,
    ppt_set_table_geometry,
    ppt_space_shapes,
    ppt_style_table,
    ppt_sync_component,
    ppt_sync_layout,
    ppt_sync_slide_chrome,
    ppt_update_stepper,
)
from powerpoint_mcp.tools.inspection import (
    ppt_analyze_containers,
    ppt_analyze_slide_structure,
    ppt_compare_slides,
    ppt_inspect_components,
    ppt_inspect_presentation,
    ppt_inspect_shape,
    ppt_inspect_slide,
    ppt_inspect_table,
    ppt_inspect_text,
    ppt_validate_slide,
    ppt_validate_slides,
)

from powerpoint_mcp.tools.rendering import (
    ppt_render_presentation,
    ppt_render_slide,
    ppt_render_slides,
    ppt_visual_diff,
)
from powerpoint_mcp.tools.versioning import (
    Session,
    get_current_session,
    get_session_manager,
    open_presentation,
    revert_session,
    save_as,
    save_session,
)
from powerpoint_mcp.utils.logging import get_logger

logger = get_logger("powerpoint_mcp.server")

# Initialize the MCPServer instance
app = MCPServer(
    name="powerpoint-mcp",
    title="PowerPoint MCP Server",
    description="Deterministic PowerPoint (.pptx) inspection, editing, rendering, and validation MCP server for Antigravity.",
    version="0.1.0",
)


# =============================================================================
# 1. Session & Lifecycle Tools
# =============================================================================

@app.tool(
    name="ppt_open",
    description=(
        "Open a PowerPoint presentation, initialize an isolated editing session with a working copy, "
        "and return session status and presentation overview."
    ),
)
def ppt_open(presentation_path: str) -> Dict[str, Any]:
    """Open a PowerPoint presentation and initialize an isolated working session.

    Args:
        presentation_path: Absolute or relative path to the .pptx presentation file.

    Returns:
        Dictionary with session_id, working_path, slide_count, dimensions, and titles.
    """
    try:
        session: Session = open_presentation(presentation_path)
        prs_info = ppt_inspect_presentation(session.working_path)

        return {
            "success": True,
            "session_id": session.session_id,
            "source_path": session.source_path,
            "working_path": session.working_path,
            "slide_count": session.slide_count,
            "dimensions": {
                "width_inches": prs_info.get("width_inches", 13.333),
                "height_inches": prs_info.get("height_inches", 7.5),
            },
            "theme": prs_info.get("theme", "Office Theme"),
            "titles": prs_info.get("titles", []),
            "layouts": prs_info.get("layouts", []),
        }
    except Exception as exc:
        return {
            "success": False,
            "error_type": type(exc).__name__,
            "message": str(exc),
            "details": {"path": presentation_path},
        }


@app.tool(
    name="ppt_save",
    description=(
        "Save session working copy back to the original presentation path, creating an automatic "
        "timestamped backup before writing."
    ),
)
def ppt_save(presentation_path: Optional[str] = None) -> Dict[str, Any]:
    """Save working copy changes back to the original presentation file with pre-save backup.

    Args:
        presentation_path: Optional override destination path (defaults to session source path).

    Returns:
        Dictionary confirming saved path and backup path.
    """
    try:
        res = save_session(destination_path=presentation_path)
        return {
            "success": True,
            "session_id": res.get("session_id"),
            "saved_path": res.get("saved_path"),
            "backup_path": res.get("backup_path"),
            "backup_created": res.get("backup_path"),
            "timestamp": res.get("timestamp"),
        }
    except Exception as exc:
        return {
            "success": False,
            "error_type": type(exc).__name__,
            "message": str(exc),
            "details": {},
        }


@app.tool(
    name="ppt_save_as",
    description=(
        "Save session working copy to a specified new destination file path without modifying the original."
    ),
)
def ppt_save_as(
    output_path: str,
    overwrite: bool = False,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Save session working copy to a new output path.

    Args:
        output_path: Target destination file path for the .pptx copy.
        overwrite: If True, overwrites destination if it already exists (creates backup first).
        presentation_path: Source presentation path if not using active session.

    Returns:
        Dictionary confirming saved path and timestamp.
    """
    try:
        res = save_as(output_path=output_path, overwrite=overwrite)
        return {
            "success": True,
            "session_id": res.get("session_id"),
            "saved_path": res.get("saved_path"),
            "backup_created": res.get("backup_path"),
            "timestamp": res.get("timestamp"),
        }
    except Exception as exc:
        return {
            "success": False,
            "error_type": type(exc).__name__,
            "message": str(exc),
            "details": {"output_path": output_path},
        }


@app.tool(
    name="ppt_revert",
    description=(
        "Discard current uncommitted edits and revert working copy to original state or a specified backup timestamp."
    ),
)
def ppt_revert(
    target: str = "original",
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Revert working copy to original file or specified backup snapshot.

    Args:
        target: 'original' or backup timestamp string / filename.
        presentation_path: Optional presentation path.

    Returns:
        Dictionary confirming revert target and working path.
    """
    try:
        res = revert_session(backup_path=target)
        return {
            "success": True,
            "session_id": res.get("session_id"),
            "reverted_to": res.get("reverted_to"),
            "working_path": res.get("working_path"),
            "timestamp": res.get("timestamp"),
        }
    except Exception as exc:
        return {
            "success": False,
            "error_type": type(exc).__name__,
            "message": str(exc),
            "details": {"target": target},
        }


# =============================================================================
# 2. Inspection Tools
# =============================================================================

@app.tool(
    name="ppt_inspect_presentation",
    description=(
        "Inspect high-level presentation metadata, slide count, dimensions, master layout names, "
        "and titles without modifying any state."
    ),
)
def tool_inspect_presentation(
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect presentation dimensions, slide count, theme, layout names, and titles.

    Args:
        presentation_path: Optional presentation path. Defaults to active session working copy.
    """
    return ppt_inspect_presentation(presentation_path=presentation_path)


@app.tool(
    name="ppt_inspect_slide",
    description=(
        "Inspect shapes on a specific slide with filtering and detail control. Defaults to concise agent-friendly summary. "
        "Supports text_only=True, include_geometry, include_style, include_images, shape_types, and semantic_roles filters."
    ),
)
def tool_inspect_slide(
    slide_number: int,
    presentation_path: Optional[str] = None,
    detail: str = "summary",
    text_only: bool = False,
    include_geometry: bool = True,
    include_style: bool = True,
    include_xml: bool = False,
    include_images: bool = True,
    shape_types: Optional[List[str]] = None,
    semantic_roles: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Inspect slide shape tree with rich filtering and detail control.

    Args:
        slide_number: 1-indexed slide number.
        presentation_path: Presentation path (defaults to active session).
        detail: 'summary' (default, concise representation) or 'full' (exhaustive shape tree).
        text_only: If True, returns only text-bearing shapes.
        include_geometry: Whether to include coordinates/dimensions.
        include_style: Whether to include font and color styling.
        include_xml: Whether to include raw XML snippets (only in full mode).
        include_images: Whether to include image/picture shapes.
        shape_types: Optional shape types filter (e.g. ['auto_shape', 'text_box']).
        semantic_roles: Optional semantic roles filter (e.g. ['title', 'body']).
    """
    return ppt_inspect_slide(
        slide_number=slide_number,
        presentation_path=presentation_path,
        detail=detail,
        text_only=text_only,
        include_geometry=include_geometry,
        include_style=include_style,
        include_xml=include_xml,
        include_images=include_images,
        shape_types=shape_types,
        semantic_roles=semantic_roles,
    )


@app.tool(
    name="ppt_inspect_text",
    description=(
        "Efficiently inspect all text-bearing shapes on a slide without full shape-tree overhead. "
        "Returns concise shape IDs, semantic roles, text content, font styling, and coordinates."
    ),
)
def tool_inspect_text(
    slide_number: int,
    include_geometry: bool = True,
    include_style: bool = True,
    include_paragraph_metadata: bool = False,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect all text-bearing shapes on a slide.

    Args:
        slide_number: 1-indexed slide number.
        include_geometry: Include bounding coordinates and dimensions (default True).
        include_style: Include typography details (font, size, weight, color) (default True).
        include_paragraph_metadata: Include per-paragraph level and bullet metadata (default False).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_inspect_text(
        slide_number=slide_number,
        include_geometry=include_geometry,
        include_style=include_style,
        include_paragraph_metadata=include_paragraph_metadata,
        presentation_path=presentation_path,
    )



@app.tool(
    name="ppt_inspect_shape",
    description=(
        "Get exhaustive details for a single shape on a slide, including all text frames, "
        "paragraphs, runs, margins, line styling, fill properties, and OOXML snippet."
    ),
)
def tool_inspect_shape(
    slide_number: int,
    shape_id: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Get deep inspection properties for a single shape by ID.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: ID of the shape to inspect.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_inspect_shape(
        slide_number=slide_number,
        shape_id=shape_id,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_inspect_components",
    description=(
        "Inspect high-level semantic visual components (headers, footers, steppers, cards, card lists, "
        "content containers) on a slide without returning thousands of lines of raw shape data."
    ),
)
def tool_inspect_components(
    slide_number: int,
    detail: str = "summary",
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect slide components, bounding boxes, properties, and constituent shape IDs.

    Args:
        slide_number: 1-indexed slide number.
        detail: 'summary' (default concise component overview) or 'full' (with child shape details).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_inspect_components(
        slide_number=slide_number,
        detail=detail,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_compare_slides",
    description=(
        "Compare geometric, typographic, and semantic layout properties between slides. "
        "Supports comparing multiple target slides against a reference slide (v1.2) or two slides (v1.1)."
    ),
)
def tool_compare_slides(
    slide_a: Optional[int] = None,
    slide_b: Optional[int] = None,
    reference_slide: Optional[int] = None,
    target_slides: Optional[List[int]] = None,
    aspects: Optional[List[str]] = None,
    match_shapes_flag: bool = True,
    render_diff: bool = False,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Compare layout and typography between slides.

    Args:
        slide_a: Legacy 1-indexed reference slide number.
        slide_b: Legacy 1-indexed target slide number.
        reference_slide: 1-indexed reference slide number for multi-slide comparison.
        target_slides: List of 1-indexed target slide numbers.
        aspects: List of comparison aspects (['components', 'geometry', 'typography', 'colors', 'spacing']).
        match_shapes_flag: Perform semantic shape matching (for 2-slide mode).
        render_diff: Perform image visual diffing if renderers available.
        presentation_path: Presentation path.
    """
    return ppt_compare_slides(
        slide_a=slide_a,
        slide_b=slide_b,
        reference_slide=reference_slide,
        target_slides=target_slides,
        aspects=aspects,
        match_shapes_flag=match_shapes_flag,
        render_diff=render_diff,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_validate_slide",
    description=(
        "Run rule-based geometric and typographic validation on a slide, detecting overlaps (VAL-01), "
        "boundary clipping (VAL-02), off-slide elements (VAL-03), text overflow (VAL-04), tiny fonts (VAL-05), "
        "and duplicate objects (VAL-07). Returns structured summary counts and issue descriptions."
    ),
)
def tool_validate_slide(
    slide_number: int,
    rules: Optional[List[str]] = None,
    presentation_path: Optional[str] = None,
    detail: str = "summary",
) -> Dict[str, Any]:
    """Validate slide geometry and text for layout defects.

    Args:
        slide_number: 1-indexed slide number.
        rules: Optional list of rule IDs to check (e.g. ['VAL-01', 'VAL-02']). Defaults to all rules.
        presentation_path: Presentation path (defaults to active session).
        detail: 'summary' (default, concise report) or 'full' (deep details dictionary).
    """
    return ppt_validate_slide(
        slide_number=slide_number,
        rules=rules,
        presentation_path=presentation_path,
        detail=detail,
    )


@app.tool(
    name="ppt_analyze_slide_structure",
    description=(
        "Analyze the complete semantic layout hierarchy, shape roles (slide_title, subtitle, card, card_title, "
        "metric, badge, body, bullet, footer, icon, image, connector), confidence scores, and logical container nesting."
    ),
)
def tool_analyze_slide_structure(
    slide_number: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Analyze slide semantic roles, confidence, and container hierarchy.

    Args:
        slide_number: 1-indexed slide number.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_analyze_slide_structure(
        slide_number=slide_number,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_analyze_containers",
    description=(
        "Identify logical containers (cards, group boxes) and their contained child elements on a slide "
        "with bounding boxes, relative child positions, and container confidence scores."
    ),
)
def tool_analyze_containers(
    slide_number: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Identify logical containers/cards and their child shapes on a slide.

    Args:
        slide_number: 1-indexed slide number.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_analyze_containers(
        slide_number=slide_number,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_inspect_table",
    description=(
        "Inspect a PowerPoint table at the cell level. Returns grid dimensions, column widths, row heights, "
        "bounding box, and cell content/formatting in compact or full diagnostic mode."
    ),
)
def tool_inspect_table(
    slide_number: int,
    table_shape_id: Optional[int] = None,
    detail: str = "compact",
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect a PowerPoint table at the cell level.

    Args:
        slide_number: 1-indexed slide number.
        table_shape_id: Optional ID of specific table shape. If omitted, finds first table on slide.
        detail: 'compact' (default grid overview) or 'full' (cell-by-cell attributes).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_inspect_table(
        slide_number=slide_number,
        table_shape_id=table_shape_id,
        detail=detail,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_validate_slides",
    description=(
        "Run validation across multiple slides (or the entire presentation), returning compact per-slide summaries "
        "of layout defects, overlaps, boundary violations, and table overflow issues."
    ),
)
def tool_validate_slides(
    slide_numbers: Optional[List[int]] = None,
    detail: bool = False,
    rules: Optional[List[str]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Validate multiple slides in batch.

    Args:
        slide_numbers: List of 1-indexed slide numbers. If omitted, validates all slides.
        detail: Return compact counts (False) or full finding lists (True).
        rules: Optional list of rule IDs to check (e.g. ['VAL-01', 'TABLE-01']).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_validate_slides(
        slide_numbers=slide_numbers,
        detail=detail,
        rules=rules,
        presentation_path=presentation_path,
    )



# =============================================================================
# 3. Editing & Mutation Tools
# =============================================================================

@app.tool(
    name="ppt_modify_shape",
    description=(
        "Deterministically update a shape's coordinates, dimensions, rotation, z-order, or apply "
        "multi-shape alignment and distribution. Only provided properties are modified."
    ),
)
def tool_modify_shape(
    slide_number: int,
    shape_id: int,
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
    align: Optional[str] = None,
    distribute: Optional[str] = None,
    target_shape_ids: Optional[List[int]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Modify shape coordinates, dimensions, rotation, z-order, alignment, or distribution.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: ID of the primary target shape.
        x: Absolute X position in inches.
        y: Absolute Y position in inches.
        width: Absolute width in inches. Must be > 0.
        height: Absolute height in inches. Must be > 0.
        rotation: Absolute rotation in degrees (0-360).
        z_order: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward', or integer.
        dx: Relative delta X in inches.
        dy: Relative delta Y in inches.
        dwidth: Relative delta width in inches.
        dheight: Relative delta height in inches.
        drotation: Relative delta rotation in degrees.
        align: Alignment mode ('left', 'center', 'right', 'top', 'middle', 'bottom').
        distribute: Distribution mode ('horizontal', 'vertical').
        target_shape_ids: Additional shape IDs for align/distribute operations.
        presentation_path: Presentation path.
    """
    return ppt_modify_shape(
        slide_number=slide_number,
        shape_id=shape_id,
        x=x,
        y=y,
        width=width,
        height=height,
        rotation=rotation,
        z_order=z_order,
        dx=dx,
        dy=dy,
        dwidth=dwidth,
        dheight=dheight,
        drotation=drotation,
        align=align,
        distribute=distribute,
        target_shape_ids=target_shape_ids,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_modify_text",
    description=(
        "Update text content and formatting in a text frame. Supports replacing full text or individual runs "
        "while strictly preserving surrounding rich-text typography (font, size, colors, weights). "
        "Supports absolute font sizing, relative deltas (font_size_delta), scaling (font_size_scale), and bounds (min_font_size, max_font_size)."
    ),
)
def tool_modify_text(
    slide_number: int,
    shape_id: int,
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
    alignment: Optional[str] = None,
    paragraph_spacing: Optional[float] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    line_spacing: Optional[float] = None,
    margins: Optional[Dict[str, float]] = None,
    paragraph_index: Optional[int] = None,
    run_index: Optional[int] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Modify text content, typography, colors, and margins with run-level style preservation.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID.
        text: New text string.
        font_family / font_name: Font name (e.g. 'Calibri', 'Arial', 'Aptos').
        font_size / font_size_pt: Absolute font point size.
        font_size_delta: Relative point delta (+2, -2) to adjust current font size.
        font_size_scale: Proportionally scale current font size by factor (e.g. 1.15).
        min_font_size / min_pt: Lower bound font size clamp in points.
        max_font_size / max_pt: Upper bound font size clamp in points.
        bold: Bold flag.
        italic: Italic flag.
        underline: Underline flag.
        color / color_rgb: RGB color hex (e.g. '#1F497D').
        alignment: Text alignment ('left', 'center', 'right', 'justify').
        paragraph_spacing / space_before: Space before paragraph in points.
        space_after: Space after paragraph in points.
        line_spacing: Line spacing in points.
        margins: Margin dict in inches: {'left': 0.1, 'top': 0.05, 'right': 0.1, 'bottom': 0.05}.
        paragraph_index: Optional 0-indexed paragraph.
        run_index: Optional 0-indexed run.
        presentation_path: Presentation path.
    """
    return ppt_modify_text(
        slide_number=slide_number,
        shape_id=shape_id,
        text=text,
        font_family=font_family,
        font_name=font_name,
        font_size=font_size,
        font_size_pt=font_size_pt,
        font_size_delta=font_size_delta,
        font_size_scale=font_size_scale,
        min_font_size=min_font_size,
        max_font_size=max_font_size,
        min_pt=min_pt,
        max_pt=max_pt,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
        color_rgb=color_rgb,
        alignment=alignment,
        paragraph_spacing=paragraph_spacing,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
        margins=margins,
        paragraph_index=paragraph_index,
        run_index=run_index,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_scale_slide_typography",
    description=(
        "Proportionally scale or adjust font sizes across all text-bearing shapes on an entire slide "
        "while strictly preserving typography hierarchy and paragraph/bullet formatting."
    ),
)
def tool_scale_slide_typography(
    slide_number: int,
    scale_factor: float = 1.0,
    font_size_delta: Optional[float] = None,
    min_pt: Optional[float] = None,
    max_pt: Optional[float] = None,
    min_font_size: Optional[float] = None,
    max_font_size: Optional[float] = None,
    include_shape_ids: Optional[List[int]] = None,
    exclude_shape_ids: Optional[List[int]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Proportionally scale or shift typography across all text-bearing shapes on a slide.

    Args:
        slide_number: 1-indexed slide number.
        scale_factor: Scale factor multiplier for font sizes (e.g. 1.15 for +15%, 0.85 for -15%).
        font_size_delta: Point shift added to font sizes (e.g. +2.0, -1.0).
        min_pt / min_font_size: Minimum resulting font size clamp in points.
        max_pt / max_font_size: Maximum resulting font size clamp in points.
        include_shape_ids: Optional list of shape IDs to exclusively scale.
        exclude_shape_ids: Optional list of shape IDs to skip.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_scale_slide_typography(
        slide_number=slide_number,
        scale_factor=scale_factor,
        font_size_delta=font_size_delta,
        min_pt=min_pt,
        max_pt=max_pt,
        min_font_size=min_font_size,
        max_font_size=max_font_size,
        include_shape_ids=include_shape_ids,
        exclude_shape_ids=exclude_shape_ids,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_copy_shape",
    description=(
        "Clone an existing shape with all formatting, fills, lines, and text styles preserved onto "
        "the same or a different slide."
    ),
)
def tool_copy_shape(
    slide_number: int,
    shape_id: int,
    target_slide_number: Optional[int] = None,
    x_offset: float = 0.2,
    y_offset: float = 0.2,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Duplicate shape with formatting onto same or target slide.

    Args:
        slide_number: 1-indexed source slide number.
        shape_id: ID of shape to clone.
        target_slide_number: Target slide number (defaults to same slide).
        x_offset: X offset in inches (default 0.2).
        y_offset: Y offset in inches (default 0.2).
        presentation_path: Presentation path.
    """
    return ppt_copy_shape(
        slide_number=slide_number,
        shape_id=shape_id,
        target_slide_number=target_slide_number,
        x_offset=x_offset,
        y_offset=y_offset,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_move_shape",
    description=(
        "Move a shape by specifying absolute coordinates (x, y) or relative deltas (dx, dy) in inches."
    ),
)
def tool_move_shape(
    slide_number: int,
    shape_id: int,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Move shape by absolute coordinates or relative offsets.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID.
        dx: Delta X shift in inches.
        dy: Delta Y shift in inches.
        x: Absolute X position in inches.
        y: Absolute Y position in inches.
        presentation_path: Presentation path.
    """
    return ppt_move_shape(
        slide_number=slide_number,
        shape_id=shape_id,
        dx=dx,
        dy=dy,
        x=x,
        y=y,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_resize_shape",
    description=(
        "Resize a shape using absolute width/height in inches or scaling multipliers with optional aspect ratio lock."
    ),
)
def tool_resize_shape(
    slide_number: int,
    shape_id: int,
    width: Optional[float] = None,
    height: Optional[float] = None,
    scale_width: Optional[float] = None,
    scale_height: Optional[float] = None,
    scale_x: Optional[float] = None,
    scale_y: Optional[float] = None,
    lock_aspect_ratio: Optional[bool] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Resize shape using absolute dimensions or scaling multipliers.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID.
        width: Absolute width in inches. Must be > 0.
        height: Absolute height in inches. Must be > 0.
        scale_width / scale_x: Width scale multiplier (e.g. 1.2 = +20%).
        scale_height / scale_y: Height scale multiplier.
        lock_aspect_ratio: Maintain aspect ratio when scaling.
        presentation_path: Presentation path.
    """
    return ppt_resize_shape(
        slide_number=slide_number,
        shape_id=shape_id,
        width=width,
        height=height,
        scale_width=scale_width,
        scale_height=scale_height,
        scale_x=scale_x,
        scale_y=scale_y,
        lock_aspect_ratio=lock_aspect_ratio,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_delete_shape",
    description="Delete a shape cleanly from a slide.",
)
def tool_delete_shape(
    slide_number: int,
    shape_id: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Delete a shape cleanly from slide shape collection.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: ID of shape to delete.
        presentation_path: Presentation path.
    """
    return ppt_delete_shape(
        slide_number=slide_number,
        shape_id=shape_id,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_batch_modify_text",
    description=(
        "Modify multiple text shapes on a slide in a single transaction with pre-validation. "
        "Preserves existing paragraph bullets and indent structures by default."
    ),
)
def tool_batch_modify_text(
    slide_number: int,
    operations: List[Dict[str, Any]],
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Batch modify text content and typography across multiple shapes on a slide.

    Args:
        slide_number: 1-indexed slide number.
        operations: List of operation dicts containing shape_id, text, font_size, font_family, bold, italic, color, alignment, etc.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_batch_modify_text(
        slide_number=slide_number,
        operations=operations,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_batch_modify_shapes",
    description=(
        "Modify multiple shape geometries (positions, sizes, rotations, z-orders) on a slide in a single transaction. "
        "Pre-validates all shape IDs and applies atomic geometry changes."
    ),
)
def tool_batch_modify_shapes(
    slide_number: int,
    operations: List[Dict[str, Any]],
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Batch modify shape geometry coordinates and dimensions across multiple shapes on a slide.

    Args:
        slide_number: 1-indexed slide number.
        operations: List of operation dicts containing shape_id and geometry changes (x, y, width, height, dx, dy, dwidth, dheight, rotation, z_order).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_batch_modify_shapes(
        slide_number=slide_number,
        operations=operations,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_align_shapes",
    description=(
        "Align multiple shapes along a common edge or center line (left, center, right, top, middle, bottom) "
        "without manual coordinate math."
    ),
)
def tool_align_shapes(
    slide_number: int,
    shape_ids: List[int],
    alignment: str,
    reference_shape_id: Optional[int] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Align shapes along an axis or to a reference shape.

    Args:
        slide_number: 1-indexed slide number.
        shape_ids: List of shape IDs to align (minimum 2).
        alignment: Alignment mode ('left', 'center', 'right', 'top', 'middle', 'bottom').
        reference_shape_id: Optional reference shape to align against.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_align_shapes(
        slide_number=slide_number,
        shape_ids=shape_ids,
        alignment=alignment,
        reference_shape_id=reference_shape_id,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_distribute_shapes",
    description=(
        "Distribute 3 or more shapes evenly across a horizontal or vertical axis using equal gaps or equal centers."
    ),
)
def tool_distribute_shapes(
    slide_number: int,
    shape_ids: List[int],
    direction: str = "horizontal",
    spacing_mode: str = "equal_gaps",
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Distribute shapes evenly across a slide axis.

    Args:
        slide_number: 1-indexed slide number.
        shape_ids: List of shape IDs to distribute (minimum 3).
        direction: 'horizontal' or 'vertical'.
        spacing_mode: 'equal_gaps' or 'equal_centers'.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_distribute_shapes(
        slide_number=slide_number,
        shape_ids=shape_ids,
        direction=direction,
        spacing_mode=spacing_mode,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_space_shapes",
    description=(
        "Set an exact fixed gap in inches between consecutive adjacent shapes horizontally or vertically."
    ),
)
def tool_space_shapes(
    slide_number: int,
    shape_ids: List[int],
    gap_inches: float,
    direction: str = "horizontal",
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Set exact fixed spacing gap between shapes.

    Args:
        slide_number: 1-indexed slide number.
        shape_ids: List of shape IDs to space (minimum 2).
        gap_inches: Gap distance in inches between adjacent shape boundaries.
        direction: 'horizontal' or 'vertical'.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_space_shapes(
        slide_number=slide_number,
        shape_ids=shape_ids,
        gap_inches=gap_inches,
        direction=direction,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_equalize_sizes",
    description=(
        "Equalize width, height, or both across multiple shapes deterministically using first, max, min, avg, or target dimensions."
    ),
)
def tool_equalize_sizes(
    slide_number: int,
    shape_ids: List[int],
    equalize_width: bool = True,
    equalize_height: bool = True,
    target_width: Optional[float] = None,
    target_height: Optional[float] = None,
    mode: str = "first",
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Equalize widths and heights across multiple shapes.

    Args:
        slide_number: 1-indexed slide number.
        shape_ids: List of shape IDs to equalize.
        equalize_width: Whether to equalize widths (default True).
        equalize_height: Whether to equalize heights (default True).
        target_width: Explicit target width in inches.
        target_height: Explicit target height in inches.
        mode: Sizing strategy ('first', 'max', 'min', 'avg').
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_equalize_sizes(
        slide_number=slide_number,
        shape_ids=shape_ids,
        equalize_width=equalize_width,
        equalize_height=equalize_height,
        target_width=target_width,
        target_height=target_height,
        mode=mode,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_move_container",
    description=(
        "Move a logical container (card) and all its nested child shapes atomically without breaking alignment or child offsets."
    ),
)
def tool_move_container(
    slide_number: int,
    container_id: int,
    x: Optional[float] = None,
    y: Optional[float] = None,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Move container and nested children atomically.

    Args:
        slide_number: 1-indexed slide number.
        container_id: ID of container/card shape.
        x: Absolute destination X coordinate in inches.
        y: Absolute destination Y coordinate in inches.
        dx: Relative delta shift X in inches.
        dy: Relative delta shift Y in inches.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_move_container(
        slide_number=slide_number,
        container_id=container_id,
        x=x,
        y=y,
        dx=dx,
        dy=dy,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_resize_container",
    description=(
        "Resize a logical container (card) and proportionally adjust/reflow contained children to maintain padding and prevent overflow."
    ),
)
def tool_resize_container(
    slide_number: int,
    container_id: int,
    width: Optional[float] = None,
    height: Optional[float] = None,
    dwidth: Optional[float] = None,
    dheight: Optional[float] = None,
    scale_width: Optional[float] = None,
    scale_height: Optional[float] = None,
    reflow_children: bool = True,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Resize container with proportional child scaling.

    Args:
        slide_number: 1-indexed slide number.
        container_id: ID of container/card shape.
        width: Absolute width in inches.
        height: Absolute height in inches.
        dwidth: Relative delta width in inches.
        dheight: Relative delta height in inches.
        scale_width: Width multiplier.
        scale_height: Height multiplier.
        reflow_children: Whether to proportionally adjust children (default True).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_resize_container(
        slide_number=slide_number,
        container_id=container_id,
        width=width,
        height=height,
        dwidth=dwidth,
        dheight=dheight,
        scale_width=scale_width,
        scale_height=scale_height,
        reflow_children=reflow_children,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_reflow_container",
    description=(
        "Deterministically stack and organize all child elements vertically inside a container with clean padding and even vertical gaps."
    ),
)
def tool_reflow_container(
    slide_number: int,
    container_id: int,
    padding_inches: float = 0.2,
    item_spacing_inches: float = 0.15,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Stack and reflow elements inside a container.

    Args:
        slide_number: 1-indexed slide number.
        container_id: ID of container/card shape.
        padding_inches: Margin padding inside container edges in inches (default 0.2).
        item_spacing_inches: Vertical spacing gap between items in inches (default 0.15).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_reflow_container(
        slide_number=slide_number,
        container_id=container_id,
        padding_inches=padding_inches,
        item_spacing_inches=item_spacing_inches,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_apply_style",
    description=(
        "Apply standard design style presets (card_default, card_accent, badge_neutral, badge_success, "
        "badge_warning, badge_danger, title_hero, title_section, metric_kpi) or transfer fill/border/font "
        "styles directly from a source shape without changing text content."
    ),
)
def tool_apply_style(
    slide_number: int,
    shape_id: Optional[int] = None,
    shape_ids: Optional[List[int]] = None,
    source_shape_id: Optional[int] = None,
    source_slide_number: Optional[int] = None,
    preset: Optional[str] = None,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    line_width_pt: Optional[float] = None,
    font_family: Optional[str] = None,
    font_size_pt: Optional[float] = None,
    font_color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Apply style preset or transfer styling between shapes.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID (or pass list in shape_ids).
        shape_ids: List of target shape IDs to style in batch.
        source_shape_id: Shape ID to copy style from.
        source_slide_number: Slide number of source shape (defaults to current slide).
        preset: Standard preset name ('card_default', 'card_accent', 'badge_neutral', 'badge_success', 'badge_warning', 'badge_danger', 'title_hero', 'title_section', 'metric_kpi').
        fill_color: Fill hex color (e.g. '#F8FAFC').
        line_color: Border line hex color (e.g. '#E2E8F0').
        line_width_pt: Border width in points.
        font_family: Font family name.
        font_size_pt: Font size in points.
        font_color: Font text hex color.
        bold: Bold flag.
        italic: Italic flag.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_apply_style(
        slide_number=slide_number,
        shape_id=shape_id,
        shape_ids=shape_ids,
        source_shape_id=source_shape_id,
        source_slide_number=source_slide_number,
        preset=preset,
        fill_color=fill_color,
        line_color=line_color,
        line_width_pt=line_width_pt,
        font_family=font_family,
        font_size_pt=font_size_pt,
        font_color=font_color,
        bold=bold,
        italic=italic,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_create_flow_diagram",
    description=(
        "Create a clean multi-step flow or process diagram with aligned nodes, connecting arrows, "
        "and typography presets without manual coordinate arithmetic."
    ),
)
def tool_create_flow_diagram(
    slide_number: int,
    steps: List[Union[str, Dict[str, Any]]],
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
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a structured flow diagram with nodes and arrows.

    Args:
        slide_number: 1-indexed slide number.
        steps: List of step strings (e.g. ['Intake', 'Processing', 'Validation']) or dicts with title, description.
        direction: 'horizontal' (default) or 'vertical'.
        shape_type: 'rounded_rectangle', 'rectangle', 'chevron', 'oval'.
        start_x: Origin X position in inches (default 1.0).
        start_y: Origin Y position in inches (default 2.2).
        total_width: Total span width in inches.
        total_height: Total span height in inches.
        node_width: Explicit node width in inches.
        node_height: Explicit node height in inches.
        node_gap: Gap distance between nodes in inches.
        style_preset: Preset style name ('card_default', 'card_accent', 'badge_primary').
        connector_style: 'arrow' (default), 'line', or 'none'.
        connector_color: Hex color for connectors (default '#94A3B8').
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_create_flow_diagram(
        slide_number=slide_number,
        steps=steps,
        direction=direction,
        shape_type=shape_type,
        start_x=start_x,
        start_y=start_y,
        total_width=total_width,
        total_height=total_height,
        node_width=node_width,
        node_height=node_height,
        node_gap=node_gap,
        style_preset=style_preset,
        connector_style=connector_style,
        connector_color=connector_color,
        presentation_path=presentation_path,
    )



@app.tool(
    name="ppt_modify_ooxml",
    description=(
        "Controlled low-level OOXML manipulation helper for features like gradients, transparency, "
        "drop shadows, and raw XML manipulation. Creates safety backup before mutation."
    ),
)
def tool_modify_ooxml(
    slide_number: int,
    shape_id: Optional[int] = None,
    operation: str = "set_attribute",
    xpath: Optional[str] = None,
    attributes: Optional[Dict[str, str]] = None,
    xml_fragment: Optional[str] = None,
    transparency_percent: Optional[float] = None,
    gradient_start: Optional[str] = None,
    gradient_end: Optional[str] = None,
    gradient_angle: float = 90.0,
    shadow_blur_pt: Optional[float] = None,
    shadow_dist_pt: Optional[float] = None,
    shadow_color: Optional[str] = None,
    shadow_alpha: Optional[float] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Perform controlled OOXML modifications.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Optional target shape ID.
        operation: 'set_attribute', 'insert_element', 'replace_element', 'transparency', 'gradient', 'shadow'.
        xpath: Target XPath expression.
        attributes: Attribute key-value dictionary.
        xml_fragment: XML snippet to insert or replace.
        transparency_percent: Transparency percentage (0..100).
        gradient_start: Gradient start color hex.
        gradient_end: Gradient end color hex.
        gradient_angle: Gradient angle in degrees.
        shadow_blur_pt: Drop shadow blur radius in points.
        shadow_dist_pt: Drop shadow distance in points.
        shadow_color: Drop shadow color hex.
        shadow_alpha: Drop shadow opacity percentage.
        presentation_path: Presentation path.
    """
    return ppt_modify_ooxml(
        slide_number=slide_number,
        shape_id=shape_id,
        operation=operation,
        xpath=xpath,
        attributes=attributes,
        xml_fragment=xml_fragment,
        transparency_percent=transparency_percent,
        gradient_start=gradient_start,
        gradient_end=gradient_end,
        gradient_angle=gradient_angle,
        shadow_blur_pt=shadow_blur_pt,
        shadow_dist_pt=shadow_dist_pt,
        shadow_color=shadow_color,
        shadow_alpha=shadow_alpha,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_create_stepper",
    description=(
        "Create a clean multi-step process stepper/breadcrumb component with active/inactive step styling "
        "and connecting arrows without manual coordinate arithmetic."
    ),
)
def tool_create_stepper(
    slide_number: int,
    steps: List[str],
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
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a stepper/breadcrumb component on a slide.

    Args:
        slide_number: 1-indexed slide number.
        steps: List of step labels (e.g. ['ANALYZE', 'CONNECT', 'CONFIGURE', 'RUN']).
        active_step: Currently active step label.
        start_x: Origin X in inches.
        start_y: Origin Y in inches.
        total_width: Total span width in inches.
        node_height: Height of each step node in inches.
        shape_type: Node shape type ('rounded_rectangle', 'rectangle', 'chevron').
        active_fill: Fill color hex for active step.
        inactive_fill: Fill color hex for inactive steps.
        active_text_color: Text color hex for active step.
        inactive_text_color: Text color hex for inactive steps.
        connector_color: Color hex for connecting arrows.
        font_family: Font family name.
        font_size_pt: Font size in points.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_create_stepper(
        slide_number=slide_number,
        steps=steps,
        active_step=active_step,
        start_x=start_x,
        start_y=start_y,
        total_width=total_width,
        node_height=node_height,
        shape_type=shape_type,
        active_fill=active_fill,
        inactive_fill=inactive_fill,
        active_text_color=active_text_color,
        inactive_text_color=inactive_text_color,
        connector_color=connector_color,
        font_family=font_family,
        font_size_pt=font_size_pt,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_update_stepper",
    description=(
        "Update the active step highlight on an existing stepper component. Atomically cleans up "
        "all old stepper shapes so no orphaned background rectangles or connectors remain."
    ),
)
def tool_update_stepper(
    slide_number: int,
    active_step: str,
    steps: Optional[List[str]] = None,
    active_fill: Optional[str] = None,
    inactive_fill: Optional[str] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Update active step highlighting on an existing stepper component.

    Args:
        slide_number: 1-indexed slide number.
        active_step: Label of the step to activate (e.g. 'CONNECT').
        steps: Optional override list of step labels.
        active_fill: Optional active step fill color hex.
        inactive_fill: Optional inactive step fill color hex.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_update_stepper(
        slide_number=slide_number,
        active_step=active_step,
        steps=steps,
        active_fill=active_fill,
        inactive_fill=inactive_fill,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_sync_component",
    description=(
        "Synchronize a visual component (header, footer, stepper, content_area, card_list) from a source "
        "slide to target slides, copying geometry, styling, and structure while preserving target content."
    ),
)
def tool_sync_component(
    source_slide: int,
    source_component: str,
    target_slides: List[int],
    preserve_content: bool = True,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Synchronize a visual component across slides.

    Args:
        source_slide: 1-indexed reference slide number.
        source_component: Identifier or type of source component ('header', 'footer', 'stepper', 'card_list', 'content_area').
        target_slides: List of 1-indexed target slide numbers.
        preserve_content: Whether to preserve target slide-specific text (default True).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_sync_component(
        source_slide=source_slide,
        source_component=source_component,
        target_slides=target_slides,
        preserve_content=preserve_content,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_sync_slide_chrome",
    description=(
        "Synchronize shared slide chrome (headers, footers, steppers, margins, title treatments) "
        "across slides from a reference slide while preserving slide-specific substantive body content."
    ),
)
def tool_sync_slide_chrome(
    reference_slide: int,
    target_slides: List[int],
    components: Optional[List[str]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Synchronize shared slide chrome across target slides.

    Args:
        reference_slide: 1-indexed reference slide number.
        target_slides: List of 1-indexed target slide numbers.
        components: Optional list of components to synchronize (default: ['header', 'footer', 'stepper', 'title_treatment', 'margins']).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_sync_slide_chrome(
        reference_slide=reference_slide,
        target_slides=target_slides,
        components=components,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_sync_layout",
    description=(
        "Synchronize content area / container layout from reference slide to target slides (position, "
        "dimensions, internal spacing, typography, borders, fills) while preserving target text."
    ),
)
def tool_sync_layout(
    reference_slide: int,
    target_slides: List[int],
    component: str = "content_area",
    preserve_content: bool = True,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Synchronize layout structure across slides.

    Args:
        reference_slide: 1-indexed reference slide number.
        target_slides: List of 1-indexed target slide numbers.
        component: Target component to sync (default: 'content_area' or 'card_list').
        preserve_content: Whether to preserve target text (default True).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_sync_layout(
        reference_slide=reference_slide,
        target_slides=target_slides,
        component=component,
        preserve_content=preserve_content,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_create_structured_card_list",
    description=(
        "Create a composite structured card container with organized item rows, descriptions, and "
        "optional dividers matching standard design presets."
    ),
)
def tool_create_structured_card_list(
    slide_number: int,
    container_bbox: Dict[str, float],
    items: List[Dict[str, Any]],
    divider: bool = True,
    style_preset: str = "card_default",
    container_fill: Optional[str] = None,
    container_border: Optional[str] = None,
    title_color: Optional[str] = None,
    desc_color: Optional[str] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a structured card list with rows and dividers.

    Args:
        slide_number: 1-indexed slide number.
        container_bbox: Dictionary with 'left', 'top', 'width', 'height' in inches.
        items: List of item dictionaries with 'title' and optional 'description'.
        divider: Whether to insert dividing lines between rows (default True).
        style_preset: Preset style name ('card_default', 'card_accent').
        container_fill: Override background fill hex color.
        container_border: Override border stroke hex color.
        title_color: Override item title font color hex.
        desc_color: Override item description font color hex.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_create_structured_card_list(
        slide_number=slide_number,
        container_bbox=container_bbox,
        items=items,
        divider=divider,
        style_preset=style_preset,
        container_fill=container_fill,
        container_border=container_border,
        title_color=title_color,
        desc_color=desc_color,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_move_component",
    description=(
        "Move a logical visual component (header, stepper, card_1, content_area) and all its "
        "constituent shapes atomically preserving relative internal alignment and offsets."
    ),
)
def tool_move_component(
    slide_number: int,
    component_id: str,
    x: Optional[float] = None,
    y: Optional[float] = None,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Move component shapes atomically.

    Args:
        slide_number: 1-indexed slide number.
        component_id: Identifier or type of target component ('header', 'stepper', 'card_1', 'content_area').
        x: Absolute destination X coordinate in inches for component top-left.
        y: Absolute destination Y coordinate in inches for component top-left.
        dx: Relative delta shift X in inches.
        dy: Relative delta shift Y in inches.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_move_component(
        slide_number=slide_number,
        component_id=component_id,
        x=x,
        y=y,
        dx=dx,
        dy=dy,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_resize_component",
    description=(
        "Resize a logical visual component (content_area, card_1, card_list) and proportionally "
        "adjust/reflow all its constituent shapes atomically."
    ),
)
def tool_resize_component(
    slide_number: int,
    component_id: str,
    width: Optional[float] = None,
    height: Optional[float] = None,
    dwidth: Optional[float] = None,
    dheight: Optional[float] = None,
    scale_width: Optional[float] = None,
    scale_height: Optional[float] = None,
    reflow_children: bool = True,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Resize component and constituent shapes atomically.

    Args:
        slide_number: 1-indexed slide number.
        component_id: Identifier or type of target component ('content_area', 'card_1', 'card_list').
        width: Absolute target width in inches.
        height: Absolute target height in inches.
        dwidth: Relative delta width in inches.
        dheight: Relative delta height in inches.
        scale_width: Width scale multiplier (e.g. 1.15).
        scale_height: Height scale multiplier (e.g. 1.10).
        reflow_children: Whether to proportionally adjust child shape positions and sizes (default True).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_resize_component(
        slide_number=slide_number,
        component_id=component_id,
        width=width,
        height=height,
        dwidth=dwidth,
        dheight=dheight,
        scale_width=scale_width,
        scale_height=scale_height,
        reflow_children=reflow_children,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_add_picture",
    description=(
        "Insert an image (PNG, JPEG, BMP) onto a slide with exact coordinates and aspect ratio handling. "
        "Calculates dimensions automatically if width or height is omitted."
    ),
)
def tool_add_picture(
    slide_number: int,
    image_path: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    preserve_aspect_ratio: bool = True,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Insert an image onto a slide.

    Args:
        slide_number: 1-indexed slide number.
        image_path: Path to the image file.
        left: Left coordinate in inches.
        top: Top coordinate in inches.
        width: Target width in inches (optional).
        height: Target height in inches (optional).
        preserve_aspect_ratio: Maintain native aspect ratio (default True).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_add_picture(
        slide_number=slide_number,
        image_path=image_path,
        left=left,
        top=top,
        width=width,
        height=height,
        preserve_aspect_ratio=preserve_aspect_ratio,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_replace_picture",
    description=(
        "Replace an existing picture shape or placeholder rectangle with a new image, preserving original "
        "geometry (position, size, and rotation)."
    ),
)
def tool_replace_picture(
    slide_number: int,
    shape_id: int,
    image_path: str,
    preserve_geometry: bool = True,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Replace an existing picture or placeholder with a new image.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: ID of the existing shape or placeholder.
        image_path: Path to the new replacement image.
        preserve_geometry: Retain original bounds and rotation (default True).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_replace_picture(
        slide_number=slide_number,
        shape_id=shape_id,
        image_path=image_path,
        preserve_geometry=preserve_geometry,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_batch_modify_table_cells",
    description=(
        "Batch modify multiple cells in a PowerPoint table atomically. Supports updating cell text, font size, "
        "bold, italic, font family, font color, cell background fill, and horizontal/vertical alignment."
    ),
)
def tool_batch_modify_table_cells(
    slide_number: int,
    table_shape_id: int,
    mutations: List[Dict[str, Any]],
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Batch modify table cells.

    Args:
        slide_number: 1-indexed slide number.
        table_shape_id: ID of the table shape.
        mutations: List of cell mutation dicts with row, column, text, and optional styles.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_batch_modify_table_cells(
        slide_number=slide_number,
        table_shape_id=table_shape_id,
        mutations=mutations,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_set_table_geometry",
    description=(
        "Modify table bounding box, column widths, or row heights with PATCH semantics. "
        "Only explicitly provided properties are altered."
    ),
)
def tool_set_table_geometry(
    slide_number: int,
    table_shape_id: int,
    left: Optional[float] = None,
    top: Optional[float] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    column_widths: Optional[List[float]] = None,
    row_heights: Optional[List[float]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Modify table geometry and column/row dimensions.

    Args:
        slide_number: 1-indexed slide number.
        table_shape_id: ID of the table shape.
        left: Left position in inches.
        top: Top position in inches.
        width: Total width in inches.
        height: Total height in inches.
        column_widths: List of individual column widths in inches.
        row_heights: List of individual row heights in inches.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_set_table_geometry(
        slide_number=slide_number,
        table_shape_id=table_shape_id,
        left=left,
        top=top,
        width=width,
        height=height,
        column_widths=column_widths,
        row_heights=row_heights,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_style_table",
    description=(
        "Apply formatting (fill, fonts, alignment, margins, borders) to an entire table or specific cell range "
        "(e.g. 'all', 'row:0', 'col:1', '0:0-1:2')."
    ),
)
def tool_style_table(
    slide_number: int,
    table_shape_id: int,
    range: Optional[str] = None,
    style: Optional[Dict[str, Any]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Apply styling to a table or range of cells.

    Args:
        slide_number: 1-indexed slide number.
        table_shape_id: ID of the table shape.
        range: Range selector ('all', '0:0', '0:0-1:2', 'row:0', 'col:1').
        style: Dictionary with styling properties (fill, font_name, font_size, bold, italic, font_color, alignment, margins, borders).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_style_table(
        slide_number=slide_number,
        table_shape_id=table_shape_id,
        range=range,
        style=style,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_merge_table_cells",
    description="Merge a rectangular block of cells in a PowerPoint table.",
)
def tool_merge_table_cells(
    slide_number: int,
    table_shape_id: int,
    start_row: int,
    start_column: int,
    end_row: int,
    end_column: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Merge cells in a table.

    Args:
        slide_number: 1-indexed slide number.
        table_shape_id: ID of the table shape.
        start_row: Top row index (0-indexed).
        start_column: Left column index (0-indexed).
        end_row: Bottom row index (0-indexed).
        end_column: Right column index (0-indexed).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_merge_table_cells(
        slide_number=slide_number,
        table_shape_id=table_shape_id,
        start_row=start_row,
        start_column=start_column,
        end_row=end_row,
        end_column=end_column,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_batch_modify_tables",
    description=(
        "Execute multi-table mutations across one or multiple slides in a single atomic transaction. "
        "Supports cells, geometry, styles, and merges."
    ),
)
def tool_batch_modify_tables(
    operations: List[Dict[str, Any]],
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Batch modify multiple tables across slides.

    Args:
        operations: List of table operation dicts (each specifying slide, table, cells, geometry, styles, merge).
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_batch_modify_tables(
        operations=operations,
        presentation_path=presentation_path,
    )



# =============================================================================
# 4. Rendering & Visual Verification Tools
# =============================================================================

@app.tool(
    name="ppt_render_slide",
    description=(
        "Render a single slide to high-resolution PNG using PowerPoint COM (Windows) or LibreOffice headless."
    ),
)
def tool_render_slide(
    slide_number: int,
    output_dir: Optional[str] = None,
    output_path: Optional[str] = None,
    renderer: str = "auto",
    dpi: int = 150,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Render a single slide to PNG.

    Args:
        slide_number: 1-indexed slide number.
        output_dir: Destination directory for PNG.
        output_path: Direct output PNG file path.
        renderer: Preferred renderer engine ('auto', 'powerpoint', 'libreoffice', 'mock').
        dpi: Render resolution DPI (default 150).
        presentation_path: Presentation path.
    """
    return ppt_render_slide(
        slide_number=slide_number,
        output_dir=output_dir,
        output_path=output_path,
        renderer=renderer,
        dpi=dpi,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_render_slides",
    description=(
        "Render a specified batch of slides (e.g. [3, 4, 5, 6]) to high-resolution PNGs in one call."
    ),
)
def tool_render_slides(
    slide_numbers: List[int],
    output_dir: Optional[str] = None,
    renderer: str = "auto",
    dpi: int = 150,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Render a specified list of slides in batch to PNG.

    Args:
        slide_numbers: List of 1-indexed slide numbers to render.
        output_dir: Output directory path.
        renderer: Preferred renderer engine ('auto', 'powerpoint', 'libreoffice', 'mock').
        dpi: Render resolution DPI (default 150).
        presentation_path: Presentation path.
    """
    return ppt_render_slides(
        slide_numbers=slide_numbers,
        output_dir=output_dir,
        renderer=renderer,
        dpi=dpi,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_render_presentation",
    description="Render all slides in the presentation to PNG images.",
)
def tool_render_presentation(
    output_dir: Optional[str] = None,
    renderer: str = "auto",
    dpi: int = 150,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Render all slides in presentation to PNG images.

    Args:
        output_dir: Output directory path.
        renderer: Renderer preference ('auto', 'powerpoint', 'libreoffice', 'mock').
        dpi: Resolution DPI (default 150).
        presentation_path: Presentation path.
    """
    return ppt_render_presentation(
        output_dir=output_dir,
        renderer=renderer,
        dpi=dpi,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_visual_diff",
    description=(
        "Perform deterministic pixel-level image comparison between two slide renders, "
        "generating a diff image, changed bounding boxes, and similarity metrics."
    ),
)
def tool_visual_diff(
    before_image: str,
    after_image: str,
    output_diff_path: Optional[str] = None,
    threshold: float = 0.1,
) -> Dict[str, Any]:
    """Compute pixel difference, changed regions, and similarity percentage between two images.

    Args:
        before_image: Path to baseline PNG image.
        after_image: Path to comparison PNG image.
        output_diff_path: Optional destination path for diff heatmap overlay image.
        threshold: Pixel difference sensitivity threshold (default 0.1 / 25).
    """
    return ppt_visual_diff(
        before_image=before_image,
        after_image=after_image,
        output_diff_path=output_diff_path,
        threshold=threshold,
    )


# =============================================================================
# 5. MCP Resources
# =============================================================================

@app.resource(
    "ppt://current/presentation",
    mime_type="application/json",
    description="Presentation metadata summary, slide count, dimensions, and titles index.",
)
def resource_current_presentation() -> str:
    """Read presentation overview resource."""
    mgr = get_session_manager()
    session = mgr.get_current_session()
    if not session or not session.working_path:
        return json.dumps({
            "error": "No active presentation session",
            "message": "Please open a presentation first using ppt_open",
        })

    info = ppt_inspect_presentation(session.working_path)
    return json.dumps(info, indent=2)


@app.resource(
    "ppt://current/slide/{slide_number}",
    mime_type="application/json",
    description="Complete shape tree, coordinates, semantic roles, and typography for a slide.",
)
def resource_current_slide(slide_number: str) -> str:
    """Read structured slide shape tree resource."""
    mgr = get_session_manager()
    session = mgr.get_current_session()
    if not session or not session.working_path:
        return json.dumps({
            "error": "No active presentation session",
            "message": "Please open a presentation first using ppt_open",
        })

    try:
        s_num = int(slide_number)
        info = ppt_inspect_slide(s_num, session.working_path)
        return json.dumps(info, indent=2)
    except Exception as exc:
        return json.dumps({
            "error": "Invalid slide number or inspection failure",
            "message": str(exc),
        })


@app.resource(
    "ppt://current/slide/{slide_number}/render",
    mime_type="image/png",
    description="High-resolution PNG image render of the specified slide.",
)
def resource_current_slide_render(slide_number: str) -> bytes:
    """Read rendered PNG binary image resource for a slide."""
    mgr = get_session_manager()
    session = mgr.get_current_session()
    if not session or not session.working_path:
        raise ValueError("No active presentation session found. Call ppt_open first.")

    s_num = int(slide_number)
    res = ppt_render_slide(s_num, presentation_path=session.working_path)
    if not res.get("success") or not res.get("image_path"):
        raise RuntimeError(f"Rendering slide {s_num} failed: {res.get('message')}")

    img_file = Path(res["image_path"])
    if not img_file.exists():
        raise FileNotFoundError(f"Rendered image file missing at {img_file}")

    with open(img_file, "rb") as f:
        return f.read()


# =============================================================================
# 6. Main Entry Point
# =============================================================================

def main() -> None:
    """Main CLI entrypoint running stdio transport."""
    logger.info("Starting PowerPoint MCP Server on stdio transport...")
    app.run(transport="stdio")


if __name__ == "__main__":
    main()
