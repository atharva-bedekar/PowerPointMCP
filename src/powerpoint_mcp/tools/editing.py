"""Editing tools for PowerPoint MCP server."""

from functools import wraps
import os
from pathlib import Path
import traceback
from typing import Any, Dict, List, Optional, Union

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt

from powerpoint_mcp.models.shape import (
    AlignmentType,
    DistributionMode,
    SpacingMode,
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.pptx.editor import (
    copy_shape,
    delete_shape,
    modify_shape,
    modify_text,
    move_shape,
    resize_shape,
)
from powerpoint_mcp.pptx.geometry import align_shapes, distribute_shapes
from powerpoint_mcp.pptx.ooxml import (
    NAMESPACES,
    get_raw_shape_xml,
    safe_modify_xml,
    set_drop_shadow,
    set_gradient_fill,
    set_shape_transparency,
)
from powerpoint_mcp.tools.inspection import handle_tool_errors
from powerpoint_mcp.tools.versioning import get_session_manager


def _get_target_presentation(
    presentation_path: Optional[str] = None, operation: str = "edit"
) -> Any:
    """Resolve presentation file and ensure pre-mutation backup is made.

    Returns:
        tuple of (resolved_path_str, Presentation_instance, session_or_None)
    """
    mgr = get_session_manager()
    session = mgr.get_current_session()

    if presentation_path:
        target_p = Path(presentation_path).resolve()
        if not target_p.exists():
            raise FileNotFoundError(f"Presentation not found: {target_p}")
        target_path_str = str(target_p)
    elif session and session.working_path and Path(session.working_path).exists():
        target_path_str = str(Path(session.working_path).resolve())
    else:
        raise ValueError("No presentation path provided and no active editing session found. Please call ppt_open first.")

    # Create safety backup before mutation
    try:
        if session and (
            Path(target_path_str) == Path(session.working_path).resolve()
            or Path(target_path_str) == Path(session.source_path).resolve()
        ):
            mgr.create_backup(
                session.session_id,
                operation=operation,
                label=f"Pre-mutation backup ({operation})",
            )
        else:
            mgr.create_backup(target_path_str, operation=operation)
    except Exception:
        pass

    prs = Presentation(target_path_str)
    return target_path_str, prs, session


# Simple helper type for annotation
Tuple_Target = Any


@handle_tool_errors
def ppt_modify_shape(
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
        shape_id: ID of primary target shape.
        x: Absolute X coordinate in inches.
        y: Absolute Y coordinate in inches.
        width: Absolute width in inches. Must be > 0.
        height: Absolute height in inches. Must be > 0.
        rotation: Absolute rotation in degrees (0..360).
        z_order: Z-order index or action ('bring_to_front', 'send_to_back', 'bring_forward', 'send_backward').
        dx: Delta X shift in inches.
        dy: Delta Y shift in inches.
        dwidth: Delta width in inches.
        dheight: Delta height in inches.
        drotation: Delta rotation in degrees.
        align: Alignment mode ('left', 'center', 'right', 'top', 'middle', 'bottom').
        distribute: Distribution mode ('horizontal', 'vertical').
        target_shape_ids: Additional shape IDs to align/distribute with primary shape.
        presentation_path: Presentation path. If omitted, uses active session working copy.

    Returns:
        Structured dictionary detailing updated shape geometry.
    """
    if width is not None and width <= 0:
        raise ValueError(f"Shape width must be positive, got {width}")
    if height is not None and height <= 0:
        raise ValueError(f"Shape height must be positive, got {height}")

    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"modify_shape_{shape_id}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    slide = prs.slides[slide_number - 1]

    # Check multi-shape alignment / distribution
    if align or distribute:
        all_ids = [shape_id] + (target_shape_ids or [])
        selected_shapes = []
        for sid in all_ids:
            found = False
            for s in slide.shapes:
                if s.shape_id == sid:
                    selected_shapes.append(s)
                    found = True
                    break
            if not found:
                raise ValueError(f"Shape with ID {sid} not found on slide {slide_number}")

        if align:
            align_shapes(selected_shapes, align)
        if distribute:
            distribute_shapes(selected_shapes, distribute)

        prs.save(target_path)
        if session:
            session.save_metadata()

        return {
            "success": True,
            "shape_id": shape_id,
            "operation": "multi_shape_geometry",
            "aligned": align,
            "distributed": distribute,
            "affected_shape_ids": all_ids,
        }

    # Single shape modification
    res = modify_shape(
        slide,
        shape_id,
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
    )

    prs.save(target_path)
    if session:
        session.save_metadata()

    return {
        "success": True,
        "shape_id": shape_id,
        "updated_properties": {
            "x_inches": res["x"],
            "y_inches": res["y"],
            "width_inches": res["width"],
            "height_inches": res["height"],
            "rotation": res["rotation"],
        },
        "shape": res,
    }


@handle_tool_errors
def ppt_modify_text(
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
    """Modify text content, typography, colors, and margins while preserving surrounding styles.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID.
        text: New text string.
        font_family / font_name: Font name (e.g. 'Calibri', 'Arial', 'Aptos').
        font_size / font_size_pt: Font point size.
        bold: Bold weight flag.
        italic: Italic flag.
        underline: Underline flag.
        color / color_rgb: Hex RGB color string (e.g. '#1F497D' or '1F497D').
        alignment: Text alignment ('left', 'center', 'right', 'justify').
        paragraph_spacing / space_before: Space before paragraph in points.
        space_after: Space after paragraph in points.
        line_spacing: Line spacing in points.
        margins: Margin dict in inches: {'left': 0.1, 'top': 0.05, 'right': 0.1, 'bottom': 0.05}.
        paragraph_index: Optional 0-indexed paragraph to modify.
        run_index: Optional 0-indexed run to modify.
        presentation_path: Presentation path.

    Returns:
        Structured dictionary detailing updated text and paragraph counts.
    """
    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"modify_text_{shape_id}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    slide = prs.slides[slide_number - 1]

    res = modify_text(
        slide,
        shape_id,
        text=text,
        font_family=font_family or font_name,
        font_size=font_size if font_size is not None else font_size_pt,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color or color_rgb,
        alignment=alignment,
        paragraph_spacing=paragraph_spacing if paragraph_spacing is not None else space_before,
        space_after=space_after,
        line_spacing=line_spacing,
        margins=margins,
    )

    prs.save(target_path)
    if session:
        session.save_metadata()

    return {
        "success": True,
        "shape_id": shape_id,
        "text_summary": (res.get("text") or "")[:100],
        "paragraph_count": res.get("paragraph_count", 1),
    }


@handle_tool_errors
def ppt_copy_shape(
    slide_number: int,
    shape_id: int,
    target_slide_number: Optional[int] = None,
    x_offset: float = 0.2,
    y_offset: float = 0.2,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Clone an existing shape with all styling and relationships onto the same or a different slide.

    Args:
        slide_number: 1-indexed source slide number.
        shape_id: ID of shape to clone.
        target_slide_number: Destination slide number (defaults to same slide).
        x_offset: Horizontal offset in inches for copied shape.
        y_offset: Vertical offset in inches for copied shape.
        presentation_path: Presentation path.

    Returns:
        Structured dictionary containing new_shape_id and target_slide number.
    """
    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"copy_shape_{shape_id}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    dest_slide_num = target_slide_number if target_slide_number is not None else slide_number
    if dest_slide_num < 1 or dest_slide_num > len(prs.slides):
        raise IndexError(f"Target slide number {dest_slide_num} is out of range (1..{len(prs.slides)})")

    src_slide = prs.slides[slide_number - 1]
    dest_slide = prs.slides[dest_slide_num - 1]

    new_shape_id = copy_shape(
        src_slide,
        shape_id,
        target_slide=dest_slide,
        offset_x_inches=x_offset,
        offset_y_inches=y_offset,
    )

    prs.save(target_path)
    if session:
        session.save_metadata()

    return {
        "success": True,
        "new_shape_id": new_shape_id,
        "source_shape_id": shape_id,
        "target_slide": dest_slide_num,
    }


@handle_tool_errors
def ppt_move_shape(
    slide_number: int,
    shape_id: int,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Move a shape by absolute coordinates (x, y) or relative deltas (dx, dy) in inches.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID.
        dx: Relative delta X in inches.
        dy: Relative delta Y in inches.
        x: Absolute X coordinate in inches.
        y: Absolute Y coordinate in inches.
        presentation_path: Presentation path.

    Returns:
        Updated shape position dictionary.
    """
    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"move_shape_{shape_id}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    slide = prs.slides[slide_number - 1]
    res = move_shape(slide, shape_id, dx=dx, dy=dy, x=x, y=y)

    prs.save(target_path)
    if session:
        session.save_metadata()

    return {
        "success": True,
        "shape_id": shape_id,
        "x_inches": res["x"],
        "y_inches": res["y"],
    }


@handle_tool_errors
def ppt_resize_shape(
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
    """Resize a shape using absolute dimensions (width, height) or scale multipliers.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Target shape ID.
        width: Absolute width in inches. Must be > 0.
        height: Absolute height in inches. Must be > 0.
        scale_width / scale_x: Width scale multiplier (e.g. 1.2 = +20%).
        scale_height / scale_y: Height scale multiplier.
        lock_aspect_ratio: Maintain aspect ratio when scaling.
        presentation_path: Presentation path.

    Returns:
        Updated shape dimension dictionary.
    """
    if width is not None and width <= 0:
        raise ValueError(f"Width must be positive, got {width}")
    if height is not None and height <= 0:
        raise ValueError(f"Height must be positive, got {height}")

    eff_scale_x = scale_width if scale_width is not None else scale_x
    eff_scale_y = scale_height if scale_height is not None else scale_y

    if lock_aspect_ratio and eff_scale_x is not None and eff_scale_y is None:
        eff_scale_y = eff_scale_x
    elif lock_aspect_ratio and eff_scale_y is not None and eff_scale_x is None:
        eff_scale_x = eff_scale_y

    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"resize_shape_{shape_id}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    slide = prs.slides[slide_number - 1]
    res = resize_shape(slide, shape_id, width=width, height=height, scale_x=eff_scale_x, scale_y=eff_scale_y)

    prs.save(target_path)
    if session:
        session.save_metadata()

    return {
        "success": True,
        "shape_id": shape_id,
        "width_inches": res["width"],
        "height_inches": res["height"],
    }


@handle_tool_errors
def ppt_delete_shape(
    slide_number: int,
    shape_id: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Delete a shape cleanly from a slide.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: ID of shape to delete.
        presentation_path: Presentation path.

    Returns:
        Summary of deleted shape and remaining shape count/IDs.
    """
    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"delete_shape_{shape_id}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    slide = prs.slides[slide_number - 1]
    deleted = delete_shape(slide, shape_id)

    if not deleted:
        raise ValueError(f"Failed to delete shape {shape_id}")

    prs.save(target_path)
    if session:
        session.save_metadata()

    remaining_ids = [s.shape_id for s in slide.shapes]

    return {
        "success": True,
        "deleted_shape_id": shape_id,
        "remaining_shape_count": len(remaining_ids),
        "remaining_shape_ids": remaining_ids,
    }


@handle_tool_errors
def ppt_modify_ooxml(
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
    """Controlled low-level OOXML manipulation helper for gradients, transparency, drop shadows, and XML tags.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: Optional target shape ID (if modifying a specific shape).
        operation: Operation type ('set_attribute', 'insert_element', 'replace_element', 'transparency', 'gradient', 'shadow').
        xpath: XPath targeting the element to mutate.
        attributes: Key-value attributes to apply.
        xml_fragment: Valid XML snippet to insert or replace.
        transparency_percent: Transparency percentage (0.0 = solid, 100.0 = transparent).
        gradient_start: Hex start color for gradient fill (e.g. 'FFFFFF').
        gradient_end: Hex end color for gradient fill (e.g. '000000').
        gradient_angle: Linear gradient angle in degrees (default 90.0).
        shadow_blur_pt: Drop shadow blur radius in points.
        shadow_dist_pt: Drop shadow distance offset in points.
        shadow_color: Drop shadow color hex.
        shadow_alpha: Drop shadow opacity percentage (0..100).
        presentation_path: Presentation path.

    Returns:
        Result summary with XML snippet.
    """
    target_path, prs, session = _get_target_presentation(presentation_path, operation=f"ooxml_{operation}")

    if slide_number < 1 or slide_number > len(prs.slides):
        raise IndexError(f"Slide number {slide_number} is out of range (1..{len(prs.slides)})")

    slide = prs.slides[slide_number - 1]

    # Target shape or slide root
    target_obj = None
    if shape_id is not None:
        for s in slide.shapes:
            if s.shape_id == shape_id:
                target_obj = s
                break
        if target_obj is None:
            raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_number}")
    else:
        target_obj = slide._element

    # 1. Transparency helper
    if transparency_percent is not None or operation == "transparency":
        tp = transparency_percent if transparency_percent is not None else 50.0
        set_shape_transparency(target_obj, tp)

    # 2. Gradient helper
    elif gradient_start is not None or gradient_end is not None or operation == "gradient":
        set_gradient_fill(
            target_obj,
            start_hex=gradient_start or "FFFFFF",
            end_hex=gradient_end or "000000",
            angle_deg=gradient_angle,
        )

    # 3. Drop shadow helper
    elif (
        shadow_blur_pt is not None
        or shadow_dist_pt is not None
        or shadow_color is not None
        or operation == "shadow"
    ):
        set_drop_shadow(
            target_obj,
            blur_rad_pt=shadow_blur_pt if shadow_blur_pt is not None else 4.0,
            dist_pt=shadow_dist_pt if shadow_dist_pt is not None else 3.0,
            color_hex=shadow_color or "000000",
            alpha_percent=shadow_alpha if shadow_alpha is not None else 40.0,
        )

    # 4. Custom XML mutation via XPath
    elif xpath or attributes or xml_fragment:
        def mutate_fn(elem: etree._Element):
            target_nodes = elem.xpath(xpath, namespaces=NAMESPACES) if xpath else [elem]
            if not target_nodes:
                raise ValueError(f"XPath '{xpath}' did not match any elements")

            for node in target_nodes:
                if operation == "set_attribute" and attributes:
                    for k, v in attributes.items():
                        node.set(k, str(v))
                elif operation == "insert_element" and xml_fragment:
                    frag = etree.fromstring(xml_fragment.encode("utf-8"))
                    node.append(frag)
                elif operation == "replace_element" and xml_fragment:
                    frag = etree.fromstring(xml_fragment.encode("utf-8"))
                    p = node.getparent()
                    if p is not None:
                        idx = p.index(node)
                        p.remove(node)
                        p.insert(idx, frag)

        safe_modify_xml(target_obj, mutate_fn)

    prs.save(target_path)
    if session:
        session.save_metadata()

    snippet = get_raw_shape_xml(target_obj) if hasattr(target_obj, "_element") else ""
    return {
        "success": True,
        "operation": operation,
        "shape_id": shape_id,
        "xml_snippet": snippet[:500] if snippet else "<modified/>",
    }
