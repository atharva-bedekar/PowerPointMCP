"""Slide comparison combining geometric AST inspection, shape matching, typography analysis, and visual diffing."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.pptx.inspector import match_shapes
from powerpoint_mcp.rendering.image_diff import VisualDiffResult, visual_diff


@dataclass
class SlideComparisonResult:
    """Comprehensive comparison report between two slides."""

    slide_a_number: int
    slide_b_number: int
    geometric_match_score: float  # 0.0 to 100.0%
    overall_similarity_score: float  # 0.0 to 100.0%
    shape_count_a: int
    shape_count_b: int
    shape_count_diff: int
    dimension_match: bool
    layout_match: bool
    shape_matches: List[Dict[str, Any]] = field(default_factory=list)
    unmatched_shapes_a: List[Dict[str, Any]] = field(default_factory=list)
    unmatched_shapes_b: List[Dict[str, Any]] = field(default_factory=list)
    layout_differences: List[Dict[str, Any]] = field(default_factory=list)
    typography_differences: List[Dict[str, Any]] = field(default_factory=list)
    visual_diff: Optional[VisualDiffResult] = None
    summary: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Serialize comparison report to a structured dictionary."""
        return {
            "slide_a_number": self.slide_a_number,
            "slide_b_number": self.slide_b_number,
            "geometric_match_score": round(self.geometric_match_score, 2),
            "overall_similarity_score": round(self.overall_similarity_score, 2),
            "shape_count_a": self.shape_count_a,
            "shape_count_b": self.shape_count_b,
            "shape_count_diff": self.shape_count_diff,
            "dimension_match": self.dimension_match,
            "layout_match": self.layout_match,
            "matched_shape_count": len(self.shape_matches),
            "unmatched_a_count": len(self.unmatched_shapes_a),
            "unmatched_b_count": len(self.unmatched_shapes_b),
            "shape_matches": self.shape_matches,
            "unmatched_shapes_a": self.unmatched_shapes_a,
            "unmatched_shapes_b": self.unmatched_shapes_b,
            "layout_differences": self.layout_differences,
            "typography_differences": self.typography_differences,
            "visual_diff": self.visual_diff.to_dict() if self.visual_diff else None,
            "summary": self.summary,
        }


def compare_slides(
    slide_a_model: SlideModel,
    slide_b_model: SlideModel,
    slide_a_img_path: Optional[Union[str, Path]] = None,
    slide_b_img_path: Optional[Union[str, Path]] = None,
    diff_output_path: Optional[Union[str, Path]] = None,
) -> SlideComparisonResult:
    """Compare two slide models geometrically, structurally, typographically, and optionally visually.

    Args:
        slide_a_model: Baseline SlideModel (e.g. original slide or slide A).
        slide_b_model: Comparison SlideModel (e.g. modified slide or slide B).
        slide_a_img_path: Optional path to rendered PNG of slide A.
        slide_b_img_path: Optional path to rendered PNG of slide B.
        diff_output_path: Optional target path to write visual diff PNG.

    Returns:
        SlideComparisonResult detailing geometric scores, matched/unmatched shapes,
        layout shifts, typographical modifications, and optional visual diff metrics.
    """
    # 1. Dimension and layout match
    dimension_match = (
        slide_a_model.width_emu == slide_b_model.width_emu
        and slide_a_model.height_emu == slide_b_model.height_emu
    )
    layout_match = slide_a_model.layout_name == slide_b_model.layout_name

    layout_diffs: List[Dict[str, Any]] = []
    if not dimension_match:
        layout_diffs.append({
            "type": "dimensions_mismatch",
            "slide_a": f"{slide_a_model.width_inches}x{slide_a_model.height_inches} in",
            "slide_b": f"{slide_b_model.width_inches}x{slide_b_model.height_inches} in",
        })
    if not layout_match:
        layout_diffs.append({
            "type": "layout_name_mismatch",
            "slide_a": slide_a_model.layout_name,
            "slide_b": slide_b_model.layout_name,
        })

    # 2. Shape Matching
    matches = match_shapes(slide_a_model, slide_b_model, min_confidence=0.30)
    matched_a_ids = {m["shape_a_id"] for m in matches}
    matched_b_ids = {m["shape_b_id"] for m in matches}

    unmatched_a = [
        s.to_dict() for s in slide_a_model.shapes if s.shape_id not in matched_a_ids
    ]
    unmatched_b = [
        s.to_dict() for s in slide_b_model.shapes if s.shape_id not in matched_b_ids
    ]

    # 3. Analyze matched pairs for geometry shifts & typography differences
    typography_diffs: List[Dict[str, Any]] = []

    for match in matches:
        sa = slide_a_model.get_shape_by_id(match["shape_a_id"])
        sb = slide_b_model.get_shape_by_id(match["shape_b_id"])
        if not sa or not sb:
            continue

        # Check geometry shifts
        dx = round(sb.bounds.left_inches - sa.bounds.left_inches, 4)
        dy = round(sb.bounds.top_inches - sa.bounds.top_inches, 4)
        dw = round(sb.bounds.width_inches - sa.bounds.width_inches, 4)
        dh = round(sb.bounds.height_inches - sa.bounds.height_inches, 4)

        if abs(dx) > 0.01 or abs(dy) > 0.01 or abs(dw) > 0.01 or abs(dh) > 0.01:
            layout_diffs.append({
                "type": "shape_geometry_shift",
                "shape_a_id": sa.shape_id,
                "shape_b_id": sb.shape_id,
                "name": sa.name,
                "delta_x_inches": dx,
                "delta_y_inches": dy,
                "delta_width_inches": dw,
                "delta_height_inches": dh,
            })

        # Check text and typography
        if sa.text_frame or sb.text_frame:
            text_a = (sa.text_frame.text if sa.text_frame else "").strip()
            text_b = (sb.text_frame.text if sb.text_frame else "").strip()

            if text_a != text_b:
                typography_diffs.append({
                    "type": "text_content_change",
                    "shape_a_id": sa.shape_id,
                    "shape_b_id": sb.shape_id,
                    "shape_name": sa.name,
                    "text_a": text_a,
                    "text_b": text_b,
                })

            # Check style attributes if both have text
            if sa.text_frame and sb.text_frame:
                runs_a = sa.text_frame.paragraphs[0].runs if sa.text_frame.paragraphs else []
                runs_b = sb.text_frame.paragraphs[0].runs if sb.text_frame.paragraphs else []
                if runs_a and runs_b:
                    st_a = runs_a[0].style
                    st_b = runs_b[0].style
                    style_diffs = {}
                    if st_a.font_name and st_b.font_name and st_a.font_name != st_b.font_name:
                        style_diffs["font_name"] = {"a": st_a.font_name, "b": st_b.font_name}
                    if st_a.font_size_pt and st_b.font_size_pt and abs(st_a.font_size_pt - st_b.font_size_pt) > 0.1:
                        style_diffs["font_size_pt"] = {"a": st_a.font_size_pt, "b": st_b.font_size_pt}
                    if st_a.bold is not None and st_b.bold is not None and st_a.bold != st_b.bold:
                        style_diffs["bold"] = {"a": st_a.bold, "b": st_b.bold}
                    if st_a.italic is not None and st_b.italic is not None and st_a.italic != st_b.italic:
                        style_diffs["italic"] = {"a": st_a.italic, "b": st_b.italic}
                    if st_a.color_rgb and st_b.color_rgb and st_a.color_rgb != st_b.color_rgb:
                        style_diffs["color_rgb"] = {"a": st_a.color_rgb, "b": st_b.color_rgb}

                    if style_diffs:
                        typography_diffs.append({
                            "type": "typography_style_change",
                            "shape_a_id": sa.shape_id,
                            "shape_b_id": sb.shape_id,
                            "shape_name": sa.name,
                            "differences": style_diffs,
                        })

    # 4. Geometric Match Score calculation
    max_shape_count = max(len(slide_a_model.shapes), len(slide_b_model.shapes))
    if max_shape_count == 0:
        geometric_match_score = 100.0
    else:
        matched_sum = sum(m["confidence_score"] for m in matches)
        geometric_match_score = (matched_sum / float(max_shape_count)) * 100.0
        geometric_match_score = max(0.0, min(100.0, geometric_match_score))

    # 5. Visual Diff (if rendered images provided)
    visual_res: Optional[VisualDiffResult] = None
    if slide_a_img_path and slide_b_img_path:
        visual_res = visual_diff(
            slide_a_img_path, slide_b_img_path, diff_output_path=diff_output_path
        )
        overall_similarity_score = (
            0.5 * geometric_match_score + 0.5 * visual_res.similarity_percentage
        )
    else:
        overall_similarity_score = geometric_match_score

    overall_similarity_score = max(0.0, min(100.0, overall_similarity_score))

    # 6. Generate summary
    summary_parts = []
    if geometric_match_score >= 99.9 and not layout_diffs and not typography_diffs:
        summary_parts.append("Slides are structurally and geometrically identical (100% match).")
    else:
        summary_parts.append(
            f"Compared Slide {slide_a_model.slide_number} with Slide {slide_b_model.slide_number}: "
            f"{len(matches)} matched shapes, {len(unmatched_a)} unmatched in A, {len(unmatched_b)} unmatched in B. "
            f"Geometric match score: {geometric_match_score:.1f}%."
        )
        if layout_diffs:
            summary_parts.append(f"Detected {len(layout_diffs)} layout/geometry change(s).")
        if typography_diffs:
            summary_parts.append(f"Detected {len(typography_diffs)} text/typography change(s).")

    if visual_res:
        summary_parts.append(
            f"Visual image similarity: {visual_res.similarity_percentage:.2f}% "
            f"({visual_res.pixel_diff_count} changed pixels, {len(visual_res.changed_bounding_boxes)} changed region(s))."
        )

    summary = " ".join(summary_parts)

    return SlideComparisonResult(
        slide_a_number=slide_a_model.slide_number,
        slide_b_number=slide_b_model.slide_number,
        geometric_match_score=geometric_match_score,
        overall_similarity_score=overall_similarity_score,
        shape_count_a=len(slide_a_model.shapes),
        shape_count_b=len(slide_b_model.shapes),
        shape_count_diff=len(slide_b_model.shapes) - len(slide_a_model.shapes),
        dimension_match=dimension_match,
        layout_match=layout_match,
        shape_matches=matches,
        unmatched_shapes_a=unmatched_a,
        unmatched_shapes_b=unmatched_b,
        layout_differences=layout_diffs,
        typography_differences=typography_diffs,
        visual_diff=visual_res,
        summary=summary,
    )
