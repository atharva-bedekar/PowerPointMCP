"""Deterministic semantic role inference and container/card hierarchy analysis for PowerPoint slides."""

from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from powerpoint_mcp.models.shape import (
    BoundingBox,
    SemanticRole,
    ShapeType,
    emu_to_inches,
    emu_to_pt,
    inches_to_emu,
)
from powerpoint_mcp.pptx.editor import _resolve_slide
from powerpoint_mcp.pptx.styles import extract_fill_style, extract_line_style


def _is_metric_text(text: str) -> bool:
    """Check if text represents a numeric metric, KPI, percentage, or currency figure."""
    cleaned = text.strip()
    if not cleaned or len(cleaned) > 25:
        return False
    # Patterns like $10M, 99.9%, +45%, 10x, 4.2k, 100/100, 1.5B
    pattern = r"^[\$€£¥\+\-]?\d+([.,]\d+)?\s*(%|x|k|m|b|ms|s|fps|pts|gb|tb|pb|usd)?$"
    return bool(re.match(pattern, cleaned, re.IGNORECASE))


def _is_badge_text(text: str) -> bool:
    """Check if text represents a compact badge or status pill."""
    cleaned = text.strip()
    if not cleaned or len(cleaned) > 20:
        return False
    keywords = {"active", "beta", "new", "done", "in progress", "prod", "deprecated", "v1.0", "v1.1", "live", "passed", "failed", "critical", "warning", "high", "medium", "low"}
    if cleaned.lower() in keywords:
        return True
    return bool(cleaned.isupper() and len(cleaned) <= 15 and " " not in cleaned)


def infer_element_role_and_confidence(
    shape: Any,
    slide_w_emu: int,
    slide_h_emu: int,
    container_ids: Optional[Set[int]] = None,
) -> Tuple[SemanticRole, float]:
    """Infer deterministic semantic role and confidence score for a shape."""
    slide_w_in = emu_to_inches(slide_w_emu) if slide_w_emu > 0 else 13.333
    slide_h_in = emu_to_inches(slide_h_emu) if slide_h_emu > 0 else 7.5

    left_in = emu_to_inches(int(getattr(shape, "left", 0)))
    top_in = emu_to_inches(int(getattr(shape, "top", 0)))
    width_in = emu_to_inches(int(getattr(shape, "width", 0)))
    height_in = emu_to_inches(int(getattr(shape, "height", 0)))
    shape_name = getattr(shape, "name", "")
    shape_type_val = getattr(shape, "shape_type", None)

    # 1. Background check: shape covering >= 90% of the canvas
    if width_in >= slide_w_in * 0.90 and height_in >= slide_h_in * 0.90:
        return SemanticRole.BACKGROUND, 0.98

    # 2. Table / Chart check
    if getattr(shape, "has_table", False) or shape_type_val in (MSO_SHAPE_TYPE.TABLE, 19):
        return SemanticRole.TABLE, 0.99
    if getattr(shape, "has_chart", False) or shape_type_val in (MSO_SHAPE_TYPE.CHART, 3):
        return SemanticRole.CHART, 0.99

    # 3. Line / Connector check
    if shape_type_val in (MSO_SHAPE_TYPE.LINE, 9, MSO_SHAPE_TYPE.FREEFORM, 5):
        return SemanticRole.CONNECTOR, 0.95

    # 4. Image / Icon check
    if shape_type_val in (MSO_SHAPE_TYPE.PICTURE, 13) or hasattr(shape, "image"):
        if width_in <= 1.25 and height_in <= 1.25:
            return SemanticRole.ICON, 0.90
        return SemanticRole.IMAGE, 0.95

    # 5. Placeholder inspection
    if getattr(shape, "is_placeholder", False):
        try:
            ph_type = shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, 1, 3):
                return SemanticRole.SLIDE_TITLE, 0.98
            elif ph_type in (PP_PLACEHOLDER.SUBTITLE, 4):
                return SemanticRole.SUBTITLE, 0.92
            elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, 2, 7):
                return SemanticRole.BODY, 0.88
            elif ph_type in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.HEADER, 15, 16, 14, 10):
                return SemanticRole.FOOTER, 0.96
        except Exception:
            pass

    # 6. Text inspection
    has_text = False
    text_content = ""
    max_font_size = 14.0
    paragraphs_count = 0
    is_bold = False

    if getattr(shape, "has_text_frame", False):
        try:
            tf = shape.text_frame
            text_content = (tf.text or "").strip()
            if text_content:
                has_text = True
                paragraphs_count = len(tf.paragraphs)
                found_sizes = []
                for p in tf.paragraphs:
                    if getattr(p, "font", None) and p.font.size is not None:
                        try:
                            found_sizes.append(float(p.font.size.pt))
                        except Exception:
                            found_sizes.append(emu_to_pt(int(p.font.size)))
                    for r in p.runs:
                        if getattr(r, "font", None) and r.font.size is not None:
                            try:
                                found_sizes.append(float(r.font.size.pt))
                            except Exception:
                                found_sizes.append(emu_to_pt(int(r.font.size)))
                        if r.font and r.font.bold:
                            is_bold = True
                if found_sizes:
                    max_font_size = max(found_sizes)
        except Exception:
            pass

    norm_top = top_in / slide_h_in if slide_h_in > 0 else 0.0

    if has_text:
        # Check Footer
        if norm_top >= 0.86 or any(k in shape_name.lower() for k in ("footer", "slide number", "date", "page number")):
            return SemanticRole.FOOTER, 0.95

        # Check Metric
        if _is_metric_text(text_content) or (max_font_size >= 24 and len(text_content) <= 15 and any(c.isdigit() for c in text_content)):
            return SemanticRole.METRIC, 0.92

        # Check Badge / Pill
        if (width_in <= 2.2 and height_in <= 0.45) or _is_badge_text(text_content):
            return SemanticRole.BADGE, 0.90

        # Check Slide Title
        if (norm_top < 0.22 and max_font_size >= 22) or ("Title" in shape_name and norm_top < 0.30):
            return SemanticRole.SLIDE_TITLE, 0.95

        # Check Subtitle
        if (0.14 <= norm_top < 0.35 and 14 <= max_font_size < 22 and not container_ids) or ("Subtitle" in shape_name and norm_top < 0.40):
            return SemanticRole.SUBTITLE, 0.88

        # Check Section Header
        if "header" in shape_name.lower() or "section" in shape_name.lower() or (is_bold and max_font_size >= 16 and paragraphs_count == 1 and width_in > 3.0):
            return SemanticRole.SECTION_HEADER, 0.86

        # Check Bullet vs Body
        if paragraphs_count > 1 and any(getattr(p, "level", 0) > 0 for p in shape.text_frame.paragraphs):
            return SemanticRole.BULLET, 0.89

        return SemanticRole.BODY, 0.85

    # 7. Non-text shapes: Check for Card / Container candidate
    fill = extract_fill_style(shape)
    line = extract_line_style(shape)
    has_fill_or_line = (fill.get("type") in ("solid", "gradient", "pattern") or line.get("fill_type") == "solid")
    area_sq_in = width_in * height_in

    if has_fill_or_line and 1.5 <= area_sq_in <= (slide_w_in * slide_h_in * 0.75):
        if "card" in shape_name.lower() or "box" in shape_name.lower() or "container" in shape_name.lower():
            return SemanticRole.CARD, 0.95
        return SemanticRole.CARD, 0.88

    return SemanticRole.UNKNOWN, 0.50


def analyze_slide_structure(
    slide_or_prs: Any,
    slide_number: Optional[int] = None,
) -> Dict[str, Any]:
    """Analyze the complete semantic role hierarchy and logical containers on a slide.

    Args:
        slide_or_prs: Slide, Presentation, or presentation path.
        slide_number: 1-indexed slide number.

    Returns:
        Structured dictionary containing slide elements with roles, confidence, parents, and containers.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs

    prs_w = int(getattr(slide.part, "slide_width", 12192000)) if hasattr(slide, "part") and hasattr(slide.part, "slide_width") else 12192000
    prs_h = int(getattr(slide.part, "slide_height", 6858000)) if hasattr(slide, "part") and hasattr(slide.part, "slide_height") else 6858000
    slide_w_in = emu_to_inches(prs_w)
    slide_h_in = emu_to_inches(prs_h)

    # Pass 1: Identify all candidate container / card shapes
    containers: List[Dict[str, Any]] = []
    shapes_list = list(slide.shapes)

    for idx, shape in enumerate(shapes_list):
        w_in = emu_to_inches(int(getattr(shape, "width", 0)))
        h_in = emu_to_inches(int(getattr(shape, "height", 0)))
        area_in = w_in * h_in

        # Skip slide canvas background
        if w_in >= slide_w_in * 0.90 and h_in >= slide_h_in * 0.90:
            continue

        fill = extract_fill_style(shape)
        line = extract_line_style(shape)
        has_fill_or_line = (fill.get("type") in ("solid", "gradient", "pattern") or line.get("fill_type") == "solid")

        # Potential card container: has visible boundary and reasonable area
        if has_fill_or_line and area_in >= 1.2:
            containers.append({
                "container_id": shape.shape_id,
                "shape": shape,
                "z_order": idx,
                "bbox": {
                    "x": emu_to_inches(int(getattr(shape, "left", 0))),
                    "y": emu_to_inches(int(getattr(shape, "top", 0))),
                    "width": w_in,
                    "height": h_in,
                },
                "children_ids": [],
            })

    # Pass 2: Establish parent-child containment hierarchy
    container_ids = {c["container_id"] for c in containers}
    parent_map: Dict[int, int] = {}

    for shape in shapes_list:
        sid = shape.shape_id
        if sid in container_ids:
            continue

        l_in = emu_to_inches(int(getattr(shape, "left", 0)))
        t_in = emu_to_inches(int(getattr(shape, "top", 0)))
        w_in = emu_to_inches(int(getattr(shape, "width", 0)))
        h_in = emu_to_inches(int(getattr(shape, "height", 0)))
        r_in = l_in + w_in
        b_in = t_in + h_in

        # Find enclosing container
        for c in containers:
            cb = c["bbox"]
            tol = 0.08
            if (
                l_in >= cb["x"] - tol
                and t_in >= cb["y"] - tol
                and r_in <= cb["x"] + cb["width"] + tol
                and b_in <= cb["y"] + cb["height"] + tol
            ):
                c["children_ids"].append(sid)
                parent_map[sid] = c["container_id"]
                break

    # Pass 3: Infer roles with container context
    elements: List[Dict[str, Any]] = []

    for idx, shape in enumerate(shapes_list):
        sid = shape.shape_id
        role, conf = infer_element_role_and_confidence(shape, prs_w, prs_h, container_ids)

        parent_id = parent_map.get(sid)

        # Refine role if shape is inside a container
        if parent_id is not None:
            if role == SemanticRole.SLIDE_TITLE or role == SemanticRole.TITLE:
                role = SemanticRole.CARD_TITLE
                conf = 0.90
            elif role == SemanticRole.BODY and is_top_of_card(shape, shapes_list, parent_id):
                role = SemanticRole.CARD_TITLE
                conf = 0.85

        child_ids = []
        for c in containers:
            if c["container_id"] == sid:
                child_ids = c["children_ids"]
                role = SemanticRole.CARD
                conf = 0.92
                break

        text_preview = ""
        if getattr(shape, "has_text_frame", False) and shape.text_frame.text:
            text_preview = shape.text_frame.text.strip()[:60]

        elements.append({
            "shape_id": sid,
            "name": shape.name,
            "role": role.value if isinstance(role, SemanticRole) else str(role),
            "confidence": round(conf, 2),
            "parent_id": parent_id,
            "children_ids": child_ids,
            "bbox": {
                "x": emu_to_inches(int(getattr(shape, "left", 0))),
                "y": emu_to_inches(int(getattr(shape, "top", 0))),
                "width": emu_to_inches(int(getattr(shape, "width", 0))),
                "height": emu_to_inches(int(getattr(shape, "height", 0))),
            },
            "text_preview": text_preview,
        })

    # Clean container list for output
    formatted_containers = []
    for c in containers:
        formatted_containers.append({
            "container_id": c["container_id"],
            "name": c["shape"].name,
            "role": "card",
            "confidence": 0.92,
            "bbox": c["bbox"],
            "children_ids": c["children_ids"],
            "child_count": len(c["children_ids"]),
        })

    return {
        "success": True,
        "slide_number": getattr(slide, "slide_number", slide_number or 1),
        "total_elements": len(elements),
        "total_containers": len(formatted_containers),
        "containers": formatted_containers,
        "elements": elements,
    }


def is_top_of_card(shape: Any, shapes_list: List[Any], container_id: int) -> bool:
    """Check if text shape is positioned at the top of its parent container."""
    container_shape = next((s for s in shapes_list if s.shape_id == container_id), None)
    if not container_shape:
        return False
    c_top = emu_to_inches(int(container_shape.top))
    c_height = emu_to_inches(int(container_shape.height))
    s_top = emu_to_inches(int(shape.top))
    return (s_top - c_top) <= (c_height * 0.35)


def analyze_containers(
    slide_or_prs: Any,
    slide_number: Optional[int] = None,
) -> Dict[str, Any]:
    """Analyze and return logical container structures and their children."""
    struct = analyze_slide_structure(slide_or_prs, slide_number)
    return {
        "success": True,
        "slide_number": struct["slide_number"],
        "total_containers": struct["total_containers"],
        "containers": struct["containers"],
    }


def _find_container_and_children(slide: Any, container_id: int) -> Tuple[Any, List[Any]]:
    """Locate the container shape and all its contained child shapes."""
    container_shape = None
    all_shapes = list(slide.shapes)
    for s in all_shapes:
        if s.shape_id == container_id:
            container_shape = s
            break

    if container_shape is None:
        raise ValueError(f"Container shape with ID {container_id} not found on slide")

    c_left = int(container_shape.left)
    c_top = int(container_shape.top)
    c_right = c_left + int(container_shape.width)
    c_bottom = c_top + int(container_shape.height)

    tol = inches_to_emu(0.08)
    children = []

    for s in all_shapes:
        if s.shape_id == container_id:
            continue
        s_left = int(s.left)
        s_top = int(s.top)
        s_right = s_left + int(s.width)
        s_bottom = s_top + int(s.height)

        if (
            s_left >= c_left - tol
            and s_top >= c_top - tol
            and s_right <= c_right + tol
            and s_bottom <= c_bottom + tol
        ):
            children.append(s)

    return container_shape, children


def move_container(
    slide_or_prs: Any,
    container_id: int,
    slide_number: Optional[int] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    dx: Optional[float] = None,
    dy: Optional[float] = None,
) -> Dict[str, Any]:
    """Move a logical container and all its nested child shapes atomically.

    Args:
        slide_or_prs: Slide or Presentation.
        container_id: ID of container/card shape.
        slide_number: 1-indexed slide number.
        x: Target absolute X position in inches for the container.
        y: Target absolute Y position in inches for the container.
        dx: Relative horizontal delta in inches.
        dy: Relative vertical delta in inches.

    Returns:
        Summary detailing updated container position and children moved.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    container_shape, children = _find_container_and_children(slide, container_id)

    # Compute delta shift
    if x is not None:
        dx_emu = inches_to_emu(x) - int(container_shape.left)
    elif dx is not None:
        dx_emu = inches_to_emu(dx)
    else:
        dx_emu = 0

    if y is not None:
        dy_emu = inches_to_emu(y) - int(container_shape.top)
    elif dy is not None:
        dy_emu = inches_to_emu(dy)
    else:
        dy_emu = 0

    # Shift container
    container_shape.left = int(container_shape.left) + dx_emu
    container_shape.top = int(container_shape.top) + dy_emu

    # Shift all children
    child_records = []
    for child in children:
        child.left = int(child.left) + dx_emu
        child.top = int(child.top) + dy_emu
        child_records.append({
            "shape_id": child.shape_id,
            "name": child.name,
            "x": emu_to_inches(int(child.left)),
            "y": emu_to_inches(int(child.top)),
        })

    return {
        "success": True,
        "container_id": container_id,
        "container_name": container_shape.name,
        "x": emu_to_inches(int(container_shape.left)),
        "y": emu_to_inches(int(container_shape.top)),
        "width": emu_to_inches(int(container_shape.width)),
        "height": emu_to_inches(int(container_shape.height)),
        "dx_inches": emu_to_inches(dx_emu),
        "dy_inches": emu_to_inches(dy_emu),
        "children_moved_count": len(children),
        "children": child_records,
    }


def resize_container(
    slide_or_prs: Any,
    container_id: int,
    slide_number: Optional[int] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    dwidth: Optional[float] = None,
    dheight: Optional[float] = None,
    scale_width: Optional[float] = None,
    scale_height: Optional[float] = None,
    reflow_children: bool = True,
) -> Dict[str, Any]:
    """Resize a logical container and scale/reflow its children to preserve layout proportions.

    Args:
        slide_or_prs: Slide or Presentation.
        container_id: ID of container/card shape.
        slide_number: 1-indexed slide number.
        width: Absolute width in inches.
        height: Absolute height in inches.
        dwidth: Relative delta width in inches.
        dheight: Relative delta height in inches.
        scale_width: Width scale multiplier.
        scale_height: Height scale multiplier.
        reflow_children: Whether to proportionally adjust child coordinates and dimensions.

    Returns:
        Summary detailing resulting container and children geometry.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    container_shape, children = _find_container_and_children(slide, container_id)

    old_l = int(container_shape.left)
    old_t = int(container_shape.top)
    old_w = max(1, int(container_shape.width))
    old_h = max(1, int(container_shape.height))

    # Calculate new width
    if width is not None:
        new_w_in = width
    elif scale_width is not None:
        new_w_in = emu_to_inches(old_w) * float(scale_width)
    elif dwidth is not None:
        new_w_in = emu_to_inches(old_w) + float(dwidth)
    else:
        new_w_in = emu_to_inches(old_w)

    # Calculate new height
    if height is not None:
        new_h_in = height
    elif scale_height is not None:
        new_h_in = emu_to_inches(old_h) * float(scale_height)
    elif dheight is not None:
        new_h_in = emu_to_inches(old_h) + float(dheight)
    else:
        new_h_in = emu_to_inches(old_h)

    new_w_emu = inches_to_emu(new_w_in)
    new_h_emu = inches_to_emu(new_h_in)

    w_ratio = float(new_w_emu) / float(old_w)
    h_ratio = float(new_h_emu) / float(old_h)

    container_shape.width = new_w_emu
    container_shape.height = new_h_emu

    child_records = []
    if reflow_children:
        for child in children:
            rel_l = int(child.left) - old_l
            rel_t = int(child.top) - old_t

            child.left = old_l + int(round(rel_l * w_ratio))
            child.top = old_t + int(round(rel_t * h_ratio))
            child.width = max(inches_to_emu(0.2), int(round(int(child.width) * w_ratio)))

            # Adjust height if non-text or proportional vertical scaling
            if not getattr(child, "has_text_frame", False) or abs(h_ratio - 1.0) > 0.15:
                child.height = max(inches_to_emu(0.2), int(round(int(child.height) * h_ratio)))

            child_records.append({
                "shape_id": child.shape_id,
                "name": child.name,
                "x": emu_to_inches(int(child.left)),
                "y": emu_to_inches(int(child.top)),
                "width": emu_to_inches(int(child.width)),
                "height": emu_to_inches(int(child.height)),
            })

    return {
        "success": True,
        "container_id": container_id,
        "container_name": container_shape.name,
        "width": emu_to_inches(new_w_emu),
        "height": emu_to_inches(new_h_emu),
        "w_scale": round(w_ratio, 3),
        "h_scale": round(h_ratio, 3),
        "children_count": len(children),
        "children": child_records,
    }


def reflow_container(
    slide_or_prs: Any,
    container_id: int,
    slide_number: Optional[int] = None,
    padding_inches: float = 0.2,
    item_spacing_inches: float = 0.15,
) -> Dict[str, Any]:
    """Deterministically stack and organize all child elements inside a container.

    Args:
        slide_or_prs: Slide or Presentation.
        container_id: ID of container/card shape.
        slide_number: 1-indexed slide number.
        padding_inches: Margin padding inside container edges in inches.
        item_spacing_inches: Vertical gap between stacked child items in inches.

    Returns:
        Summary detailing reflowed child shapes and positions.
    """
    slide = _resolve_slide(slide_or_prs, slide_number) if (not hasattr(slide_or_prs, "shapes") or hasattr(slide_or_prs, "slides")) else slide_or_prs
    container_shape, children = _find_container_and_children(slide, container_id)

    if not children:
        return {
            "success": True,
            "container_id": container_id,
            "message": "Container has no child elements to reflow",
            "children_reflowed_count": 0,
        }

    # Sort children into top-to-bottom hierarchy
    def _child_rank(shape: Any) -> Tuple[int, int]:
        name = shape.name.lower()
        if "icon" in name or (shape.shape_type == MSO_SHAPE_TYPE.PICTURE and emu_to_inches(int(shape.width)) <= 1.2):
            return (1, int(shape.top))
        if "badge" in name or "pill" in name or (emu_to_inches(int(shape.height)) <= 0.4 and emu_to_inches(int(shape.width)) <= 2.2):
            return (2, int(shape.top))
        if "title" in name or (getattr(shape, "has_text_frame", False) and int(shape.top) < int(container_shape.top) + int(container_shape.height) * 0.4):
            return (3, int(shape.top))
        if "body" in name or "bullet" in name or getattr(shape, "has_text_frame", False):
            return (4, int(shape.top))
        return (5, int(shape.top))

    sorted_children = sorted(children, key=_child_rank)

    pad_emu = inches_to_emu(padding_inches)
    gap_emu = inches_to_emu(item_spacing_inches)
    c_left = int(container_shape.left)
    c_top = int(container_shape.top)
    c_w = int(container_shape.width)
    content_w = max(inches_to_emu(0.5), c_w - 2 * pad_emu)

    curr_y = c_top + pad_emu
    reflowed_records = []

    for child in sorted_children:
        child_w = int(child.width)
        child_h = int(child.height)

        is_compact = (
            emu_to_inches(child_h) <= 0.5
            and (emu_to_inches(child_w) <= 2.2 or "badge" in child.name.lower() or "icon" in child.name.lower())
        )

        child.top = curr_y
        child.left = c_left + pad_emu

        if not is_compact and getattr(child, "has_text_frame", False):
            child.width = content_w

        curr_y += int(child.height) + gap_emu

        reflowed_records.append({
            "shape_id": child.shape_id,
            "name": child.name,
            "x": emu_to_inches(int(child.left)),
            "y": emu_to_inches(int(child.top)),
            "width": emu_to_inches(int(child.width)),
            "height": emu_to_inches(int(child.height)),
        })

    return {
        "success": True,
        "container_id": container_id,
        "container_name": container_shape.name,
        "padding_inches": padding_inches,
        "item_spacing_inches": item_spacing_inches,
        "children_reflowed_count": len(reflowed_records),
        "children": reflowed_records,
    }
