"""Deterministic semantic component detection and inspection engine for PowerPoint slides."""

from __future__ import annotations

import math
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from pptx import Presentation

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
from powerpoint_mcp.pptx.editor import _resolve_slide
from powerpoint_mcp.pptx.inspector import inspect_slide
from powerpoint_mcp.pptx.structure import (
    analyze_containers,
    analyze_slide_structure,
    infer_element_role_and_confidence,
)


def _compute_enclosing_bbox(shapes: List[ShapeModel]) -> Optional[BoundingBox]:
    """Compute the combined enclosing bounding box for a group of shapes."""
    if not shapes:
        return None
    min_left = min(s.bbox.left_emu for s in shapes)
    min_top = min(s.bbox.top_emu for s in shapes)
    max_right = max(s.bbox.right_emu for s in shapes)
    max_bottom = max(s.bbox.bottom_emu for s in shapes)
    return BoundingBox(
        left_emu=min_left,
        top_emu=min_top,
        width_emu=max(1, max_right - min_left),
        height_emu=max(1, max_bottom - min_top),
    )


def _extract_shape_text(shape: ShapeModel) -> str:
    """Extract trimmed plain text from shape model."""
    if shape.text_frame and shape.text_frame.text:
        return shape.text_frame.text.strip()
    return ""


def _is_connector_shape(shape: ShapeModel) -> bool:
    """Check if shape is an arrow or connector."""
    name_lower = shape.name.lower()
    if shape.shape_type == ShapeType.CONNECTOR or shape.semantic_role == SemanticRole.CONNECTOR:
        return True
    if any(k in name_lower for k in ("arrow", "connector", "flow arrow", "chevron arrow", "line")):
        return True
    return False


def _is_stepper_shape(shape: ShapeModel) -> bool:
    """Check if shape name or text looks like a stepper step or breadcrumb node."""
    if _is_connector_shape(shape):
        return False

    # A main slide title is never a stepper node
    if shape.semantic_role in (SemanticRole.SLIDE_TITLE, SemanticRole.TITLE) or ("title" in shape.name.lower() and "card" not in shape.name.lower()):
        return False

    name_lower = shape.name.lower()
    if any(k in name_lower for k in ("flow node", "step node", "step", "breadcrumb", "chevron", "stepper step", "stage")):
        return True

    txt = _extract_shape_text(shape).upper()
    if txt and shape.bbox.height_inches <= 1.2:
        if any(w == txt or txt == f"STEP {w}" or txt.startswith("STEP ") for w in ("ANALYZE", "CONNECT", "CONFIGURE", "RUN", "BUILD", "DEPLOY", "TEST", "PLAN", "1", "2", "3", "4")):
            return True
    return False


def detect_slide_components(
    slide_or_path: Any,
    slide_number: int,
) -> List[ComponentModel]:
    """Detect high-level visual components on a slide deterministically.

    Args:
        slide_or_path: Path to presentation, Presentation instance, or SlideModel.
        slide_number: 1-indexed slide number.

    Returns:
        List of detected ComponentModel instances (header, footer, stepper, cards, content_container, etc.).
    """
    if isinstance(slide_or_path, SlideModel):
        slide_model = slide_or_path
    else:
        slide_model = inspect_slide(slide_or_path, slide_number)

    shapes = slide_model.shapes
    slide_w_in = slide_model.width_inches if slide_model.width_inches > 0 else 13.333
    slide_h_in = slide_model.height_inches if slide_model.height_inches > 0 else 7.5

    components: List[ComponentModel] = []
    claimed_shape_ids: Set[int] = set()

    # -------------------------------------------------------------------------
    # 1. Header / Title Block Detection
    # -------------------------------------------------------------------------
    header_shapes: List[ShapeModel] = []
    title_text = ""
    subtitle_text = ""

    for s in shapes:
        if _is_stepper_shape(s) or _is_connector_shape(s):
            continue

        norm_top = s.bbox.top_inches / slide_h_in if slide_h_in > 0 else 0.0
        text = _extract_shape_text(s)

        # Main title / subtitle check
        if (s.semantic_role in (SemanticRole.SLIDE_TITLE, SemanticRole.TITLE) or ("title" in s.name.lower() and norm_top < 0.25)) and s.bbox.height_inches <= 1.5:
            header_shapes.append(s)
            if not title_text and text:
                title_text = text
        elif (s.semantic_role == SemanticRole.SUBTITLE or "subtitle" in s.name.lower()) and norm_top < 0.35 and s.bbox.height_inches <= 1.2:
            header_shapes.append(s)
            if not subtitle_text and text:
                subtitle_text = text
        elif norm_top < 0.10 and s.bbox.height_inches <= 0.5:
            # Top-aligned header badge, logo, or accent shape
            if s.bbox.width_inches < slide_w_in * 0.95:  # Not background
                header_shapes.append(s)

    if header_shapes:
        for s in header_shapes:
            claimed_shape_ids.add(s.shape_id)
        bbox = _compute_enclosing_bbox(header_shapes)
        components.append(
            ComponentModel(
                component_id="header",
                component_type=ComponentType.HEADER,
                slide_number=slide_number,
                shape_ids=[s.shape_id for s in header_shapes],
                bbox=bbox,
                properties={
                    "title": title_text,
                    "subtitle": subtitle_text,
                    "shape_count": len(header_shapes),
                },
                confidence=0.96 if title_text else 0.85,
            )
        )

    # -------------------------------------------------------------------------
    # 2. Footer Detection
    # -------------------------------------------------------------------------
    footer_shapes: List[ShapeModel] = []
    footer_text = ""

    for s in shapes:
        if s.shape_id in claimed_shape_ids:
            continue
        norm_top = s.bbox.top_inches / slide_h_in if slide_h_in > 0 else 0.0
        name_lower = s.name.lower()
        if (
            norm_top >= 0.85
            or s.semantic_role == SemanticRole.FOOTER
            or any(k in name_lower for k in ("footer", "slide number", "page number", "confidential"))
        ):
            if s.bbox.width_inches < slide_w_in * 0.95:  # Not full background
                footer_shapes.append(s)
                text = _extract_shape_text(s)
                if text and not footer_text:
                    footer_text = text

    if footer_shapes:
        for s in footer_shapes:
            claimed_shape_ids.add(s.shape_id)
        bbox = _compute_enclosing_bbox(footer_shapes)
        components.append(
            ComponentModel(
                component_id="footer",
                component_type=ComponentType.FOOTER,
                slide_number=slide_number,
                shape_ids=[s.shape_id for s in footer_shapes],
                bbox=bbox,
                properties={
                    "text": footer_text,
                    "shape_count": len(footer_shapes),
                },
                confidence=0.95,
            )
        )

    # -------------------------------------------------------------------------
    # 3. Stepper / Breadcrumb / Process Flow Detection
    # -------------------------------------------------------------------------
    stepper_candidate_shapes: List[ShapeModel] = []
    stepper_connector_shapes: List[ShapeModel] = []

    for s in shapes:
        if s.shape_id in claimed_shape_ids:
            continue
        if _is_stepper_shape(s):
            stepper_candidate_shapes.append(s)
        elif _is_connector_shape(s):
            stepper_connector_shapes.append(s)

    # Also detect sequence of horizontally aligned shapes with step text if not named explicitly
    if len(stepper_candidate_shapes) < 2:
        # Check for horizontal sequence of 2-8 shapes with similar y and height
        unclaimed = [s for s in shapes if s.shape_id not in claimed_shape_ids and s.bbox.width_inches < slide_w_in * 0.90]
        horiz_groups: Dict[int, List[ShapeModel]] = {}
        for s in unclaimed:
            # Bucket by top coordinate (within 0.35 in)
            bucket = int(round(s.bbox.top_inches / 0.35))
            horiz_groups.setdefault(bucket, []).append(s)

        for bucket, grp in horiz_groups.items():
            if len(grp) >= 2:
                # Check if elements are horizontally sequenced with text
                sorted_grp = sorted(grp, key=lambda x: x.bbox.left_inches)
                has_step_labels = any(
                    any(w in _extract_shape_text(x).upper() for w in ("ANALYZE", "CONNECT", "CONFIGURE", "RUN", "STEP", "PHASE", "BUILD", "DEPLOY", "TEST", "PLAN", "1", "2", "3"))
                    for x in sorted_grp if _extract_shape_text(x)
                )
                if has_step_labels and len(sorted_grp) >= 2 and sorted_grp[0].bbox.top_inches < slide_h_in * 0.45:
                    stepper_candidate_shapes = sorted_grp
                    break

    if len(stepper_candidate_shapes) >= 2:
        # Sort candidates horizontally
        sorted_steps = sorted(stepper_candidate_shapes, key=lambda x: x.bbox.left_inches)
        step_labels: List[str] = []
        step_nodes: List[Dict[str, Any]] = []
        stepper_shape_ids: List[int] = []

        # Find any connectors or sub-shapes that overlap with the stepper zone
        stepper_box = _compute_enclosing_bbox(sorted_steps)
        if stepper_box:
            # Include associated connectors located within/near the stepper box
            for conn in stepper_connector_shapes:
                if (
                    conn.bbox.top_inches >= stepper_box.top_inches - 0.2
                    and conn.bbox.bottom_inches <= stepper_box.bottom_inches + 0.2
                ):
                    stepper_shape_ids.append(conn.shape_id)
                    claimed_shape_ids.add(conn.shape_id)

        # Analyze active step styling: look for distinct fill color or bold text
        fills = [s.fill.get("color_rgb", "") if s.fill else "" for s in sorted_steps]
        font_colors = []
        for s in sorted_steps:
            fc = ""
            if s.text_frame and s.text_frame.paragraphs:
                for p in s.text_frame.paragraphs:
                    if p.runs and p.runs[0].style.color_rgb:
                        fc = p.runs[0].style.color_rgb
                        break
            font_colors.append(fc)

        # Majority fill color is inactive, minority/distinct is active
        fill_counts: Dict[str, int] = {}
        for f in fills:
            if f:
                fill_counts[f] = fill_counts.get(f, 0) + 1
        majority_fill = max(fill_counts.items(), key=lambda item: item[1])[0] if fill_counts else ""

        active_step_label = None

        for i, s in enumerate(sorted_steps):
            stepper_shape_ids.append(s.shape_id)
            claimed_shape_ids.add(s.shape_id)
            lbl = _extract_shape_text(s) or s.name
            # Strip step prefixes if any for clean label
            clean_lbl = re.sub(r"^(flow node \d+\s*-\s*|step \d+\s*:\s*|\d+\.\s*)", "", lbl, flags=re.IGNORECASE).strip()
            step_labels.append(clean_lbl)

            s_fill = s.fill.get("color_rgb", "") if s.fill else ""
            is_active = False
            if majority_fill and s_fill and s_fill != majority_fill:
                is_active = True
            elif "active" in s.name.lower():
                is_active = True

            if is_active and active_step_label is None:
                active_step_label = clean_lbl

            step_nodes.append({
                "step_index": i,
                "label": clean_lbl,
                "shape_id": s.shape_id,
                "bbox": s.bbox.to_dict() if s.bbox else {},
                "is_active": is_active,
                "fill_color": s_fill,
            })

        # Default first step as active if no distinct active found
        if active_step_label is None and step_labels:
            active_step_label = step_labels[0]
            step_nodes[0]["is_active"] = True

        components.append(
            ComponentModel(
                component_id="stepper",
                component_type=ComponentType.STEPPER,
                slide_number=slide_number,
                shape_ids=stepper_shape_ids,
                bbox=stepper_box,
                properties={
                    "steps": step_labels,
                    "active_step": active_step_label,
                    "direction": "horizontal",
                    "step_nodes": step_nodes,
                },
                confidence=0.95,
            )
        )

    # -------------------------------------------------------------------------
    # 4. Containers / Cards & Card Lists Detection
    # -------------------------------------------------------------------------
    unclaimed_shapes = [s for s in shapes if s.shape_id not in claimed_shape_ids and s.bbox.width_inches < slide_w_in * 0.95]
    
    # Use structure analyzer to find cards / containers
    card_components: List[ComponentModel] = []
    for s in unclaimed_shapes:
        if s.semantic_role in (SemanticRole.CARD, SemanticRole.BACKGROUND) or "card" in s.name.lower() or "container" in s.name.lower() or "box" in s.name.lower():
            # Find children contained within this card
            children_shapes = []
            for other in unclaimed_shapes:
                if other.shape_id == s.shape_id:
                    continue
                # Spatial containment check
                if (
                    other.bbox.left_inches >= s.bbox.left_inches - 0.1
                    and other.bbox.top_inches >= s.bbox.top_inches - 0.1
                    and other.bbox.right_inches <= s.bbox.right_inches + 0.1
                    and other.bbox.bottom_inches <= s.bbox.bottom_inches + 0.1
                ):
                    children_shapes.append(other)

            card_shape_ids = [s.shape_id] + [c.shape_id for c in children_shapes]
            for cid in card_shape_ids:
                claimed_shape_ids.add(cid)

            # Extract card title and items
            card_title = ""
            card_items = []
            for c in children_shapes:
                txt = _extract_shape_text(c)
                if txt:
                    if not card_title and (c.semantic_role == SemanticRole.CARD_TITLE or "title" in c.name.lower() or c.bbox.top_inches < s.bbox.top_inches + 0.8):
                        card_title = txt
                    else:
                        card_items.append(txt)

            card_comp = ComponentModel(
                component_id=f"card_{len(card_components)+1}",
                component_type=ComponentType.CARD,
                slide_number=slide_number,
                shape_ids=card_shape_ids,
                bbox=s.bbox,
                properties={
                    "title": card_title,
                    "item_count": len(card_items),
                    "items": card_items,
                    "has_container_shape": True,
                },
                confidence=0.92,
            )
            card_components.append(card_comp)

    # Check if multiple cards form a card list
    if len(card_components) >= 2:
        # Group cards into card_list
        all_card_shape_ids = []
        for c in card_components:
            all_card_shape_ids.extend(c.shape_ids)
        combined_card_shapes = [s for s in shapes if s.shape_id in all_card_shape_ids]
        card_list_bbox = _compute_enclosing_bbox(combined_card_shapes)

        components.append(
            ComponentModel(
                component_id="card_list",
                component_type=ComponentType.CARD_LIST,
                slide_number=slide_number,
                shape_ids=all_card_shape_ids,
                bbox=card_list_bbox,
                properties={
                    "card_count": len(card_components),
                    "cards": [c.to_dict(detail="full") for c in card_components],
                },
                confidence=0.94,
            )
        )
    else:
        components.extend(card_components)

    # -------------------------------------------------------------------------
    # 5. Remaining Content Area / Content Container
    # -------------------------------------------------------------------------
    remaining_shapes = [s for s in shapes if s.shape_id not in claimed_shape_ids and s.bbox.width_inches < slide_w_in * 0.95]
    if remaining_shapes:
        content_box = _compute_enclosing_bbox(remaining_shapes)
        # Extract prominent titles/text from remaining body shapes
        body_texts = [_extract_shape_text(s) for s in remaining_shapes if _extract_shape_text(s)]
        content_title = body_texts[0] if body_texts else "Content Area"

        components.append(
            ComponentModel(
                component_id="content_area",
                component_type=ComponentType.CONTENT_AREA,
                slide_number=slide_number,
                shape_ids=[s.shape_id for s in remaining_shapes],
                bbox=content_box,
                properties={
                    "title": content_title,
                    "shape_count": len(remaining_shapes),
                    "texts": body_texts[:5],
                },
                confidence=0.90,
            )
        )

    return components


def inspect_components(
    presentation_path: str,
    slide_number: int,
    detail: str = "summary",
) -> Dict[str, Any]:
    """Inspect high-level components on a slide.

    Args:
        presentation_path: Path to presentation.
        slide_number: 1-indexed slide number.
        detail: 'summary' (concise component overview) or 'full' (with child shape details).

    Returns:
        Structured dictionary matching ppt_inspect_components schema.
    """
    comps = detect_slide_components(presentation_path, slide_number)

    return {
        "success": True,
        "slide_number": slide_number,
        "component_count": len(comps),
        "components": [c.to_dict(detail=detail) for c in comps],
    }
