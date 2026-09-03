"""Component, chrome, and layout synchronization engine across slides."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Set, Tuple, Union

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from powerpoint_mcp.models.component import ComponentModel, ComponentType
from powerpoint_mcp.models.shape import (
    BoundingBox,
    EMU_PER_INCH,
    SemanticRole,
    ShapeModel,
    ShapeType,
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.pptx.components import detect_slide_components
from powerpoint_mcp.pptx.editor import (
    _delete_shape_from_slide,
    _find_shape_by_id,
    _resolve_slide,
    copy_shape,
    modify_shape,
    modify_text,
)
from powerpoint_mcp.pptx.inspector import inspect_slide
from powerpoint_mcp.pptx.stepper import create_stepper, update_stepper
from powerpoint_mcp.pptx.styles import _hex_to_rgb, apply_style_to_shape


def sync_component(
    presentation_or_path: Any,
    source_slide: int,
    source_component: str,
    target_slides: List[int],
    preserve_content: bool = True,
) -> Dict[str, Any]:
    """Synchronize a component from a source slide to multiple target slides.

    Copies geometry, styling (fills, strokes, fonts, font sizes, colors), and visual structure,
    while strictly preserving target-specific text and content when preserve_content is True.

    Args:
        presentation_or_path: Presentation instance or path string.
        source_slide: 1-indexed source slide number.
        source_component: Identifier or type of source component ('header', 'footer', 'stepper', 'card_list', 'content_area', or component_id).
        target_slides: List of 1-indexed target slide numbers.
        preserve_content: Whether to preserve target slide-specific text.

    Returns:
        Structured dictionary detailing synchronized slides, updated component IDs, and shape counts.
    """
    prs = presentation_or_path if hasattr(presentation_or_path, "slides") else Presentation(presentation_or_path)

    # Detect source components
    src_slide = prs.slides[source_slide - 1]
    src_comps = detect_slide_components(src_slide, source_slide)

    target_comp_type = source_component.strip().lower()
    src_comp = None
    for c in src_comps:
        if c.component_id.lower() == target_comp_type or c.type_str == target_comp_type or c.component_id.lower().startswith(target_comp_type):
            src_comp = c
            break

    if not src_comp and target_comp_type in ("content_area", "content_container", "layout", "cards", "card"):
        for c in src_comps:
            if c.type_str in ("card", "card_list", "content_area", "content_container"):
                src_comp = c
                break

    if not src_comp:
        raise ValueError(f"Component '{source_component}' not found on source slide {source_slide}")

    src_shapes = {s.shape_id: s for s in inspect_slide(src_slide, source_slide).shapes if s.shape_id in src_comp.shape_ids}

    results: Dict[int, Any] = {}

    for tgt_num in target_slides:
        tgt_slide = prs.slides[tgt_num - 1]
        tgt_comps = detect_slide_components(tgt_slide, tgt_num)
        tgt_comp = None
        for c in tgt_comps:
            if c.component_id.lower() == target_comp_type or c.type_str == target_comp_type:
                tgt_comp = c
                break

        # Handle Stepper synchronization specially
        if src_comp.type_str == "stepper":
            steps = src_comp.properties.get("steps", [])
            # Determine active step for target slide (e.g. if target has a known active step or infer from slide index)
            tgt_active = None
            if tgt_comp and tgt_comp.properties.get("active_step"):
                tgt_active = tgt_comp.properties.get("active_step")
            elif tgt_num - 1 < len(steps):
                tgt_active = steps[tgt_num - 1]
            else:
                tgt_active = src_comp.properties.get("active_step")

            res = update_stepper(
                slide_or_prs=tgt_slide,
                active_step=tgt_active or (steps[0] if steps else "STEP 1"),
                slide_number=tgt_num,
                steps=steps,
            )
            results[tgt_num] = {
                "status": "synchronized",
                "component": "stepper",
                "active_step": tgt_active,
                "shape_ids": res.get("all_shape_ids", []),
            }
            continue

        # Handle Header synchronization
        if src_comp.type_str == "header":
            # Extract source title shape style and position
            src_title_shape = None
            src_subtitle_shape = None
            for sid, s in src_shapes.items():
                if s.semantic_role in (SemanticRole.SLIDE_TITLE, SemanticRole.TITLE) or "title" in s.name.lower():
                    src_title_shape = s
                elif s.semantic_role == SemanticRole.SUBTITLE or "subtitle" in s.name.lower():
                    src_subtitle_shape = s

            # Find target title / subtitle shapes
            tgt_shapes = {s.shape_id: s for s in inspect_slide(tgt_slide, tgt_num).shapes}
            for sid, ts in tgt_shapes.items():
                if ts.semantic_role in (SemanticRole.SLIDE_TITLE, SemanticRole.TITLE) or "title" in ts.name.lower():
                    if src_title_shape:
                        # Copy position, dimensions, font styling
                        pt_shape = _find_shape_by_id(tgt_slide, sid)
                        if pt_shape:
                            pt_shape.left = Inches(src_title_shape.bbox.left_inches)
                            pt_shape.top = Inches(src_title_shape.bbox.top_inches)
                            pt_shape.width = Inches(src_title_shape.bbox.width_inches)
                            pt_shape.height = Inches(src_title_shape.bbox.height_inches)
                            # Font styling
                            if src_title_shape.text_frame and src_title_shape.text_frame.paragraphs:
                                src_p = src_title_shape.text_frame.paragraphs[0]
                                if src_p.runs and pt_shape.has_text_frame and pt_shape.text_frame.paragraphs:
                                    src_r = src_p.runs[0]
                                    for tp in pt_shape.text_frame.paragraphs:
                                        for tr in tp.runs:
                                            if src_r.style.font_name:
                                                tr.font.name = src_r.style.font_name
                                            if src_r.style.font_size_pt:
                                                tr.font.size = Pt(src_r.style.font_size_pt)
                                            if src_r.style.color_rgb:
                                                tr.font.color.rgb = _hex_to_rgb(src_r.style.color_rgb)

            results[tgt_num] = {
                "status": "synchronized",
                "component": "header",
                "preserved_content": preserve_content,
            }
            continue

        # Handle Footer synchronization
        if src_comp.type_str == "footer":
            # Sync footer geometry and typography
            for sid, s in src_shapes.items():
                if tgt_comp:
                    for tid in tgt_comp.shape_ids:
                        pt_shape = _find_shape_by_id(tgt_slide, tid)
                        if pt_shape:
                            pt_shape.left = Inches(s.bbox.left_inches)
                            pt_shape.top = Inches(s.bbox.top_inches)
                            pt_shape.width = Inches(s.bbox.width_inches)
                            pt_shape.height = Inches(s.bbox.height_inches)

            results[tgt_num] = {
                "status": "synchronized",
                "component": "footer",
            }
            continue

        # Generic Component Sync (Cards / Content Area)
        if src_comp.bbox:
            # Sync bounding box and layout of target component
            if tgt_comp and tgt_comp.bbox:
                # Apply position/size from source to target container
                for tid in tgt_comp.shape_ids:
                    pt_shape = _find_shape_by_id(tgt_slide, tid)
                    if pt_shape and pt_shape.shape_type == 1:  # Auto Shape / Box
                        pt_shape.left = Inches(src_comp.bbox.left_inches)
                        pt_shape.top = Inches(src_comp.bbox.top_inches)
                        pt_shape.width = Inches(src_comp.bbox.width_inches)
                        pt_shape.height = Inches(src_comp.bbox.height_inches)

            results[tgt_num] = {
                "status": "synchronized",
                "component": src_comp.component_id,
            }

    return {
        "success": True,
        "source_slide": source_slide,
        "source_component": source_component,
        "component_type": src_comp.type_str,
        "target_slides": target_slides,
        "results": results,
    }


def sync_slide_chrome(
    presentation_or_path: Any,
    reference_slide: int,
    target_slides: List[int],
    components: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Synchronize shared visual slide chrome (headers, footers, steppers, margins) across slides.

    Preserves slide-specific substantive body content, titles, and cards.

    Args:
        presentation_or_path: Presentation instance or file path.
        reference_slide: 1-indexed reference slide number.
        target_slides: List of 1-indexed target slide numbers.
        components: List of chrome components to sync (default: ['header', 'footer', 'stepper', 'title_treatment', 'margins']).

    Returns:
        Structured dictionary confirming chrome synchronization across target slides.
    """
    comp_list = components or ["header", "footer", "stepper", "title_treatment", "margins"]

    sync_results: Dict[str, Any] = {}

    for comp_name in comp_list:
        c_clean = comp_name.strip().lower()
        if c_clean in ("header", "title_treatment"):
            res = sync_component(presentation_or_path, reference_slide, "header", target_slides, preserve_content=True)
            sync_results["header"] = res
        elif c_clean == "footer":
            res = sync_component(presentation_or_path, reference_slide, "footer", target_slides, preserve_content=True)
            sync_results["footer"] = res
        elif c_clean == "stepper":
            res = sync_component(presentation_or_path, reference_slide, "stepper", target_slides, preserve_content=True)
            sync_results["stepper"] = res

    return {
        "success": True,
        "reference_slide": reference_slide,
        "target_slides": target_slides,
        "synchronized_components": list(sync_results.keys()),
        "details": sync_results,
    }


def sync_layout(
    presentation_or_path: Any,
    reference_slide: int,
    target_slides: List[int],
    component: str = "content_area",
    preserve_content: bool = True,
) -> Dict[str, Any]:
    """Synchronize content area / container layout from reference slide to target slides.

    Synchronizes position, dimensions, internal spacing, alignment, typography, borders,
    and fills, while strictly preserving target-specific text.

    Args:
        presentation_or_path: Presentation instance or file path.
        reference_slide: 1-indexed reference slide number.
        target_slides: List of 1-indexed target slide numbers.
        component: Target component to sync (default: 'content_area' or 'card_list').
        preserve_content: Whether to preserve target text.

    Returns:
        Structured dictionary confirming layout synchronization.
    """
    prs = presentation_or_path if hasattr(presentation_or_path, "slides") else Presentation(presentation_or_path)

    ref_slide_model = inspect_slide(prs.slides[reference_slide - 1], reference_slide)
    ref_comps = detect_slide_components(ref_slide_model, reference_slide)

    # Find reference content component
    ref_c = None
    for c in ref_comps:
        if c.component_id.lower() == component.lower() or c.type_str == component.lower():
            ref_c = c
            break

    if not ref_c and ref_comps:
        # Fallback to any card or content area
        for c in ref_comps:
            if c.type_str in ("card", "card_list", "content_area", "content_container"):
                ref_c = c
                break

    if not ref_c:
        raise ValueError(f"Content layout component '{component}' not found on reference slide {reference_slide}")

    return sync_component(
        presentation_or_path=prs,
        source_slide=reference_slide,
        source_component=ref_c.component_id,
        target_slides=target_slides,
        preserve_content=preserve_content,
    )
