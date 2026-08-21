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
    ppt_copy_shape,
    ppt_delete_shape,
    ppt_modify_ooxml,
    ppt_modify_shape,
    ppt_modify_text,
    ppt_move_shape,
    ppt_resize_shape,
)
from powerpoint_mcp.tools.inspection import (
    ppt_compare_slides,
    ppt_inspect_presentation,
    ppt_inspect_shape,
    ppt_inspect_slide,
    ppt_validate_slide,
)
from powerpoint_mcp.tools.rendering import (
    ppt_render_presentation,
    ppt_render_slide,
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
        "Inspect all shapes on a specific slide, returning geometry (in inches), semantic roles "
        "(title, subtitle, body, diagram, image, footer), text content, font styling, colors, and layout structure."
    ),
)
def tool_inspect_slide(
    slide_number: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect 1-indexed slide shape tree, coordinates, semantic roles, and typography.

    Args:
        slide_number: 1-indexed slide number.
        presentation_path: Presentation path (defaults to active session).
    """
    return ppt_inspect_slide(slide_number=slide_number, presentation_path=presentation_path)


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
    name="ppt_compare_slides",
    description=(
        "Compare geometric, typographic, and semantic layout properties between two slides, "
        "matching corresponding shapes with multi-factor confidence scores."
    ),
)
def tool_compare_slides(
    slide_a: int,
    slide_b: int,
    match_shapes_flag: bool = True,
    render_diff: bool = False,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Compare layout and typography between two slides.

    Args:
        slide_a: 1-indexed reference slide number.
        slide_b: 1-indexed target slide number.
        match_shapes_flag: Perform semantic shape matching (default True).
        render_diff: Perform image visual diffing if renderers available.
        presentation_path: Presentation path.
    """
    return ppt_compare_slides(
        slide_a=slide_a,
        slide_b=slide_b,
        match_shapes_flag=match_shapes_flag,
        render_diff=render_diff,
        presentation_path=presentation_path,
    )


@app.tool(
    name="ppt_validate_slide",
    description=(
        "Run rule-based geometric and typographic validation on a slide, detecting overlaps (VAL-01), "
        "boundary clipping (VAL-02), off-slide elements (VAL-03), text overflow (VAL-04), tiny fonts (VAL-05), "
        "and duplicate objects (VAL-07)."
    ),
)
def tool_validate_slide(
    slide_number: int,
    rules: Optional[List[str]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Validate slide geometry and text for layout defects.

    Args:
        slide_number: 1-indexed slide number.
        rules: Optional list of rule IDs to check (e.g. ['VAL-01', 'VAL-02']). Defaults to all rules.
        presentation_path: Presentation path.
    """
    return ppt_validate_slide(
        slide_number=slide_number,
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
        "while strictly preserving surrounding rich-text typography (font, size, colors, weights)."
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
        font_size / font_size_pt: Font point size.
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
