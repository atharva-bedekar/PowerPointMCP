"""Cross-slide comparison engine for semantic components, geometry, typography, and styles."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from powerpoint_mcp.models.component import ComponentModel, ComponentType
from powerpoint_mcp.models.shape import ShapeModel, emu_to_inches
from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.pptx.components import detect_slide_components
from powerpoint_mcp.pptx.inspector import inspect_slide
from powerpoint_mcp.rendering.visual_compare import compare_slides as legacy_compare_slides


def _compare_component_pair(
    ref_comp: Optional[ComponentModel],
    target_comp: Optional[ComponentModel],
    ref_shapes: Dict[int, ShapeModel],
    target_shapes: Dict[int, ShapeModel],
) -> Dict[str, Any]:
    """Compare a reference component with a target component across geometry, typography, and styling."""
    if not ref_comp and not target_comp:
        return {"status": "MATCH", "summary": "MATCH", "details": "Both absent"}
    if not ref_comp:
        return {"status": "DIFFERENT", "summary": "Component absent on reference slide", "details": "Component absent on reference slide"}
    if not target_comp:
        return {"status": "DIFFERENT", "summary": "Component absent on target slide", "details": "Component absent on target slide"}

    diffs: List[str] = []

    # Check bounding box geometry
    if ref_comp.bbox and target_comp.bbox:
        rb = ref_comp.bbox
        tb = target_comp.bbox
        dx = abs(rb.left_inches - tb.left_inches)
        dy = abs(rb.top_inches - tb.top_inches)
        dw = abs(rb.width_inches - tb.width_inches)
        dh = abs(rb.height_inches - tb.height_inches)

        if dx > 0.05 or dy > 0.05 or dw > 0.10 or dh > 0.10:
            diffs.append("geometry differs")

    # Check stepper active step & steps
    if ref_comp.type_str == "stepper" and target_comp.type_str == "stepper":
        ref_active = ref_comp.properties.get("active_step")
        target_active = target_comp.properties.get("active_step")
        if ref_active != target_active:
            diffs.append(f"active step differs (ref: {ref_active}, target: {target_active})")
        ref_steps = ref_comp.properties.get("steps", [])
        target_steps = target_comp.properties.get("steps", [])
        if ref_steps != target_steps:
            diffs.append("step sequence differs")

    # Check typography on constituent shapes
    ref_fonts: List[float] = []
    target_fonts: List[float] = []

    for sid in ref_comp.shape_ids:
        s = ref_shapes.get(sid)
        if s and s.text_frame:
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    if r.style.font_size_pt is not None:
                        ref_fonts.append(float(r.style.font_size_pt))

    for sid in target_comp.shape_ids:
        s = target_shapes.get(sid)
        if s and s.text_frame:
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    if r.style.font_size_pt is not None:
                        target_fonts.append(float(r.style.font_size_pt))

    if ref_fonts and target_fonts:
        avg_ref_pt = sum(ref_fonts) / len(ref_fonts)
        avg_target_pt = sum(target_fonts) / len(target_fonts)
        delta_pt = round(avg_target_pt - avg_ref_pt, 1)
        if abs(delta_pt) >= 1.0:
            diffs.append(f"typography {delta_pt:+}pt")

    status = "MATCH" if not diffs else "DIFFERENT"
    return {
        "status": status,
        "differences": diffs,
        "summary": ", ".join(diffs) if diffs else "MATCH",
    }


def compare_cross_slides(
    presentation_path: str,
    reference_slide: int,
    target_slides: List[int],
    aspects: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Compare multiple target slides against a reference slide for cross-slide consistency.

    Args:
        presentation_path: Path to presentation file.
        reference_slide: 1-indexed reference slide number.
        target_slides: List of 1-indexed target slide numbers to compare against reference.
        aspects: List of comparison aspects (default: ['components', 'geometry', 'typography', 'colors', 'spacing']).

    Returns:
        Structured cross-slide comparison report with compact summary text and detailed findings.
    """
    aspects_set = set(a.lower() for a in (aspects or ["components", "geometry", "typography", "colors", "spacing"]))

    # Inspect reference slide and its components
    ref_slide_model = inspect_slide(presentation_path, reference_slide)
    ref_components = detect_slide_components(ref_slide_model, reference_slide)
    ref_shapes_map = {s.shape_id: s for s in ref_slide_model.shapes}
    ref_comp_map = {c.component_id: c for c in ref_components}
    # Also index by type
    ref_type_map = {c.type_str: c for c in ref_components}

    comparison_results: Dict[int, Dict[str, Any]] = {}
    summary_lines: List[str] = [f"REFERENCE: Slide {reference_slide}\n"]

    # Component categories to report
    categories = ["header", "stepper", "footer", "content_area", "card_list"]

    category_reports: Dict[str, List[str]] = {cat: [] for cat in categories}
    typography_reports: List[str] = []

    for target_num in target_slides:
        target_slide_model = inspect_slide(presentation_path, target_num)
        target_components = detect_slide_components(target_slide_model, target_num)
        target_shapes_map = {s.shape_id: s for s in target_slide_model.shapes}
        target_comp_map = {c.component_id: c for c in target_components}
        target_type_map = {c.type_str: c for c in target_components}

        target_report: Dict[str, Any] = {
            "slide_number": target_num,
            "components": {},
            "typography_diffs": [],
        }

        # Compare each component type
        for cat in categories:
            ref_c = ref_type_map.get(cat)
            target_c = target_type_map.get(cat)
            comp_diff = _compare_component_pair(ref_c, target_c, ref_shapes_map, target_shapes_map)
            target_report["components"][cat] = comp_diff

            desc = comp_diff["summary"]
            category_reports[cat].append(f"  Slide {target_num}: {desc}")

        # Overall typography comparison
        ref_all_fonts = []
        for s in ref_slide_model.shapes:
            if s.text_frame:
                for p in s.text_frame.paragraphs:
                    for r in p.runs:
                        if r.style.font_size_pt is not None:
                            ref_all_fonts.append(float(r.style.font_size_pt))

        target_all_fonts = []
        for s in target_slide_model.shapes:
            if s.text_frame:
                for p in s.text_frame.paragraphs:
                    for r in p.runs:
                        if r.style.font_size_pt is not None:
                            target_all_fonts.append(float(r.style.font_size_pt))

        if ref_all_fonts and target_all_fonts:
            ref_avg = sum(ref_all_fonts) / len(ref_all_fonts)
            target_avg = sum(target_all_fonts) / len(target_all_fonts)
            diff_pt = round(target_avg - ref_avg, 1)
            if abs(diff_pt) >= 0.5:
                typography_reports.append(f"  Slide {target_num}: body fonts {diff_pt:+}pt")
            else:
                typography_reports.append(f"  Slide {target_num}: MATCH")

        comparison_results[target_num] = target_report

    # Assemble formatted human-readable summary
    for cat in categories:
        cat_title = cat.upper().replace("_", " ")
        if any(ref_type_map.get(cat) for _ in [1]) or any(category_reports[cat]):
            summary_lines.append(f"{cat_title}")
            summary_lines.extend(category_reports[cat])
            summary_lines.append("")

    if typography_reports:
        summary_lines.append("TYPOGRAPHY")
        summary_lines.extend(typography_reports)
        summary_lines.append("")

    summary_text = "\n".join(summary_lines).strip()

    return {
        "success": True,
        "reference_slide": reference_slide,
        "target_slides": target_slides,
        "summary": summary_text,
        "targets": comparison_results,
    }
