"""Geometry engine for shape alignment, distribution, equalization, collision detection, and slide boundary validation."""

from typing import Any, Dict, List, Optional, Sequence, Tuple, Union
from pptx.enum.shapes import MSO_SHAPE_TYPE

from powerpoint_mcp.models.shape import (
    AlignmentType,
    BoundingBox,
    DistributionMode,
    EMU_PER_INCH,
    SemanticRole,
    ShapeModel,
    ShapeType,
    SpacingMode,
    emu_to_inches,
    inches_to_emu,
)
from powerpoint_mcp.models.slide import SlideModel


def _get_shape_bounds(shape: Any) -> Tuple[int, int, int, int]:
    """Extract (left_emu, top_emu, width_emu, height_emu) from any shape representation.

    Supports:
    - python-pptx Shape objects (has .left, .top, .width, .height)
    - ShapeModel objects (has .bbox)
    - BoundingBox objects (has .left_emu, .top_emu, .width_emu, .height_emu)
    - Dictionaries with EMU or inch keys
    """
    if isinstance(shape, BoundingBox):
        return (shape.left_emu, shape.top_emu, shape.width_emu, shape.height_emu)

    if isinstance(shape, ShapeModel):
        return (shape.bbox.left_emu, shape.bbox.top_emu, shape.bbox.width_emu, shape.bbox.height_emu)

    if hasattr(shape, "left") and hasattr(shape, "top") and hasattr(shape, "width") and hasattr(shape, "height"):
        return (int(shape.left), int(shape.top), int(shape.width), int(shape.height))

    if hasattr(shape, "bbox") and isinstance(shape.bbox, BoundingBox):
        return (shape.bbox.left_emu, shape.bbox.top_emu, shape.bbox.width_emu, shape.bbox.height_emu)

    if isinstance(shape, dict):
        if "left_emu" in shape and "top_emu" in shape and "width_emu" in shape and "height_emu" in shape:
            return (int(shape["left_emu"]), int(shape["top_emu"]), int(shape["width_emu"]), int(shape["height_emu"]))
        if "x" in shape and "y" in shape and "width" in shape and "height" in shape:
            return (
                inches_to_emu(float(shape["x"])),
                inches_to_emu(float(shape["y"])),
                inches_to_emu(float(shape["width"])),
                inches_to_emu(float(shape["height"])),
            )
        if "left_inches" in shape and "top_inches" in shape and "width_inches" in shape and "height_inches" in shape:
            return (
                inches_to_emu(float(shape["left_inches"])),
                inches_to_emu(float(shape["top_inches"])),
                inches_to_emu(float(shape["width_inches"])),
                inches_to_emu(float(shape["height_inches"])),
            )

    raise TypeError(f"Cannot extract bounding box from object of type {type(shape).__name__}")


def _set_shape_bounds(
    shape: Any,
    left_emu: Optional[int] = None,
    top_emu: Optional[int] = None,
    width_emu: Optional[int] = None,
    height_emu: Optional[int] = None,
) -> None:
    """Set coordinates and/or dimensions on any shape representation in EMU units."""
    if isinstance(shape, BoundingBox):
        if left_emu is not None:
            shape.left_emu = int(left_emu)
        if top_emu is not None:
            shape.top_emu = int(top_emu)
        if width_emu is not None:
            shape.width_emu = int(width_emu)
        if height_emu is not None:
            shape.height_emu = int(height_emu)
        return

    if isinstance(shape, ShapeModel):
        cur_l, cur_t, cur_w, cur_h = _get_shape_bounds(shape)
        shape.bbox = BoundingBox(
            left_emu=int(left_emu) if left_emu is not None else cur_l,
            top_emu=int(top_emu) if top_emu is not None else cur_t,
            width_emu=int(width_emu) if width_emu is not None else cur_w,
            height_emu=int(height_emu) if height_emu is not None else cur_h,
        )
        return

    if hasattr(shape, "left") and hasattr(shape, "top") and hasattr(shape, "width") and hasattr(shape, "height"):
        if left_emu is not None:
            shape.left = int(left_emu)
        if top_emu is not None:
            shape.top = int(top_emu)
        if width_emu is not None:
            shape.width = int(width_emu)
        if height_emu is not None:
            shape.height = int(height_emu)
        return

    if isinstance(shape, dict):
        if left_emu is not None:
            shape["left_emu"] = int(left_emu)
            shape["x"] = emu_to_inches(left_emu)
            shape["left_inches"] = emu_to_inches(left_emu)
        if top_emu is not None:
            shape["top_emu"] = int(top_emu)
            shape["y"] = emu_to_inches(top_emu)
            shape["top_inches"] = emu_to_inches(top_emu)
        if width_emu is not None:
            shape["width_emu"] = int(width_emu)
            shape["width"] = emu_to_inches(width_emu)
            shape["width_inches"] = emu_to_inches(width_emu)
        if height_emu is not None:
            shape["height_emu"] = int(height_emu)
            shape["height"] = emu_to_inches(height_emu)
            shape["height_inches"] = emu_to_inches(height_emu)
        return

    raise TypeError(f"Cannot update bounding box on object of type {type(shape).__name__}")


def _get_shape_id(shape: Any) -> int:
    """Extract shape ID from python-pptx Shape, ShapeModel, or dict."""
    if hasattr(shape, "shape_id"):
        return int(shape.shape_id)
    if isinstance(shape, dict):
        return int(shape.get("shape_id", shape.get("id", 0)))
    return 0


def _get_shape_name(shape: Any) -> str:
    """Extract shape name from python-pptx Shape, ShapeModel, or dict."""
    if hasattr(shape, "name"):
        return str(shape.name)
    if isinstance(shape, dict):
        return str(shape.get("name", "Shape"))
    return "Shape"


def align_shapes(
    shapes: Sequence[Any],
    alignment: Union[AlignmentType, str],
    reference_shape: Optional[Any] = None,
) -> List[Any]:
    """Align a sequence of shapes along a specified edge or center line.

    Supported alignments:
    - LEFT: Align all left edges to reference shape or leftmost shape.
    - CENTER: Align horizontal centers to reference shape or average center.
    - RIGHT: Align right edges to reference shape or rightmost shape.
    - TOP: Align top edges to reference shape or topmost shape.
    - MIDDLE: Align vertical centers to reference shape or average center.
    - BOTTOM: Align bottom edges to reference shape or bottommost shape.

    Args:
        shapes: Sequence of python-pptx Shapes, ShapeModels, BoundingBoxes, or coordinate dicts.
        alignment: Target alignment (AlignmentType enum or string).
        reference_shape: Optional reference shape to align against.

    Returns:
        The list of aligned shape objects.
    """
    if not shapes:
        return list(shapes)

    if isinstance(alignment, str):
        alignment_str = alignment.strip().lower()
    elif isinstance(alignment, AlignmentType):
        alignment_str = alignment.value.lower()
    else:
        raise ValueError(f"Invalid alignment: {alignment}")

    bounds_list = [_get_shape_bounds(s) for s in shapes]

    if reference_shape is not None:
        ref_l, ref_t, ref_w, ref_h = _get_shape_bounds(reference_shape)
        ref_cx = ref_l + ref_w // 2
        ref_cy = ref_t + ref_h // 2
        ref_r = ref_l + ref_w
        ref_b = ref_t + ref_h
    else:
        ref_l = min(b[0] for b in bounds_list)
        ref_t = min(b[1] for b in bounds_list)
        ref_r = max(b[0] + b[2] for b in bounds_list)
        ref_b = max(b[1] + b[3] for b in bounds_list)
        ref_cx = sum(b[0] + b[2] // 2 for b in bounds_list) // len(bounds_list)
        ref_cy = sum(b[1] + b[3] // 2 for b in bounds_list) // len(bounds_list)

    for shape in shapes:
        l, t, w, h = _get_shape_bounds(shape)
        if alignment_str == "left":
            _set_shape_bounds(shape, left_emu=ref_l)
        elif alignment_str == "center":
            _set_shape_bounds(shape, left_emu=ref_cx - w // 2)
        elif alignment_str == "right":
            _set_shape_bounds(shape, left_emu=ref_r - w)
        elif alignment_str == "top":
            _set_shape_bounds(shape, top_emu=ref_t)
        elif alignment_str == "middle":
            _set_shape_bounds(shape, top_emu=ref_cy - h // 2)
        elif alignment_str == "bottom":
            _set_shape_bounds(shape, top_emu=ref_b - h)
        else:
            raise ValueError(f"Unknown alignment mode: '{alignment_str}'. Supported: left, center, right, top, middle, bottom.")

    return list(shapes)


def distribute_shapes(
    shapes: Sequence[Any],
    mode: Union[DistributionMode, str],
    spacing: Union[SpacingMode, str] = SpacingMode.EQUAL_GAPS,
) -> List[Any]:
    """Distribute a sequence of shapes evenly along horizontal or vertical axes.

    Supports:
    - DistributionMode.HORIZONTAL with EQUAL_GAPS or EQUAL_CENTERS
    - DistributionMode.VERTICAL with EQUAL_GAPS or EQUAL_CENTERS

    Args:
        shapes: Sequence of shapes (minimum 3 required for inner redistribution).
        mode: DistributionMode.HORIZONTAL or DistributionMode.VERTICAL.
        spacing: SpacingMode.EQUAL_GAPS (default) or SpacingMode.EQUAL_CENTERS.

    Returns:
        The list of shapes sorted and positioned along the distribution axis.
    """
    if len(shapes) < 3:
        return list(shapes)

    mode_str = mode.value.lower() if isinstance(mode, DistributionMode) else str(mode).strip().lower()
    spacing_str = spacing.value.lower() if isinstance(spacing, SpacingMode) else str(spacing).strip().lower()

    if mode_str == "horizontal":
        if spacing_str == "equal_gaps":
            sorted_shapes = sorted(shapes, key=lambda s: _get_shape_bounds(s)[0])
            first_l, _, first_w, _ = _get_shape_bounds(sorted_shapes[0])
            last_l, _, last_w, _ = _get_shape_bounds(sorted_shapes[-1])

            total_span = (last_l + last_w) - first_l
            total_shape_width = sum(_get_shape_bounds(s)[2] for s in sorted_shapes)
            remaining_gap = total_span - total_shape_width
            gap_per_space = remaining_gap // (len(sorted_shapes) - 1)

            curr_left = first_l + first_w + gap_per_space
            for i in range(1, len(sorted_shapes) - 1):
                _, _, w, _ = _get_shape_bounds(sorted_shapes[i])
                _set_shape_bounds(sorted_shapes[i], left_emu=curr_left)
                curr_left += w + gap_per_space

            return sorted_shapes

        elif spacing_str == "equal_centers":
            sorted_shapes = sorted(shapes, key=lambda s: _get_shape_bounds(s)[0] + _get_shape_bounds(s)[2] // 2)
            c_first = _get_shape_bounds(sorted_shapes[0])[0] + _get_shape_bounds(sorted_shapes[0])[2] // 2
            c_last = _get_shape_bounds(sorted_shapes[-1])[0] + _get_shape_bounds(sorted_shapes[-1])[2] // 2

            step = (c_last - c_first) // (len(sorted_shapes) - 1)
            for i in range(1, len(sorted_shapes) - 1):
                target_center = c_first + i * step
                _, _, w, _ = _get_shape_bounds(sorted_shapes[i])
                _set_shape_bounds(sorted_shapes[i], left_emu=target_center - w // 2)

            return sorted_shapes
        else:
            raise ValueError(f"Unknown spacing mode: '{spacing_str}'. Supported: equal_gaps, equal_centers.")

    elif mode_str == "vertical":
        if spacing_str == "equal_gaps":
            sorted_shapes = sorted(shapes, key=lambda s: _get_shape_bounds(s)[1])
            _, first_t, _, first_h = _get_shape_bounds(sorted_shapes[0])
            _, last_t, _, last_h = _get_shape_bounds(sorted_shapes[-1])

            total_span = (last_t + last_h) - first_t
            total_shape_height = sum(_get_shape_bounds(s)[3] for s in sorted_shapes)
            remaining_gap = total_span - total_shape_height
            gap_per_space = remaining_gap // (len(sorted_shapes) - 1)

            curr_top = first_t + first_h + gap_per_space
            for i in range(1, len(sorted_shapes) - 1):
                _, _, _, h = _get_shape_bounds(sorted_shapes[i])
                _set_shape_bounds(sorted_shapes[i], top_emu=curr_top)
                curr_top += h + gap_per_space

            return sorted_shapes

        elif spacing_str == "equal_centers":
            sorted_shapes = sorted(shapes, key=lambda s: _get_shape_bounds(s)[1] + _get_shape_bounds(s)[3] // 2)
            c_first = _get_shape_bounds(sorted_shapes[0])[1] + _get_shape_bounds(sorted_shapes[0])[3] // 2
            c_last = _get_shape_bounds(sorted_shapes[-1])[1] + _get_shape_bounds(sorted_shapes[-1])[3] // 2

            step = (c_last - c_first) // (len(sorted_shapes) - 1)
            for i in range(1, len(sorted_shapes) - 1):
                target_center = c_first + i * step
                _, _, _, h = _get_shape_bounds(sorted_shapes[i])
                _set_shape_bounds(sorted_shapes[i], top_emu=target_center - h // 2)

            return sorted_shapes
        else:
            raise ValueError(f"Unknown spacing mode: '{spacing_str}'. Supported: equal_gaps, equal_centers.")

    else:
        raise ValueError(f"Unknown distribution mode: '{mode_str}'. Supported: horizontal, vertical.")


def equalize_dimensions(
    shapes: Sequence[Any],
    equalize_width: bool = True,
    equalize_height: bool = True,
    target_width_inches: Optional[float] = None,
    target_height_inches: Optional[float] = None,
    mode: str = "first",
) -> List[Any]:
    """Equalize widths and/or heights across a sequence of shapes.

    Args:
        shapes: Sequence of shapes to equalize.
        equalize_width: Whether to equalize widths.
        equalize_height: Whether to equalize heights.
        target_width_inches: Explicit target width in inches (overrides mode).
        target_height_inches: Explicit target height in inches (overrides mode).
        mode: Strategy for target dimension if not explicitly given:
              'first' (first shape's dimension), 'max', 'min', 'avg'.

    Returns:
        The list of modified shape objects.
    """
    if not shapes or (not equalize_width and not equalize_height):
        return list(shapes)

    bounds = [_get_shape_bounds(s) for s in shapes]

    target_w_emu: Optional[int] = None
    if equalize_width:
        if target_width_inches is not None:
            target_w_emu = inches_to_emu(target_width_inches)
        elif mode == "first":
            target_w_emu = bounds[0][2]
        elif mode == "max":
            target_w_emu = max(b[2] for b in bounds)
        elif mode == "min":
            target_w_emu = min(b[2] for b in bounds)
        elif mode == "avg":
            target_w_emu = sum(b[2] for b in bounds) // len(bounds)
        else:
            target_w_emu = bounds[0][2]

    target_h_emu: Optional[int] = None
    if equalize_height:
        if target_height_inches is not None:
            target_h_emu = inches_to_emu(target_height_inches)
        elif mode == "first":
            target_h_emu = bounds[0][3]
        elif mode == "max":
            target_h_emu = max(b[3] for b in bounds)
        elif mode == "min":
            target_h_emu = min(b[3] for b in bounds)
        elif mode == "avg":
            target_h_emu = sum(b[3] for b in bounds) // len(bounds)
        else:
            target_h_emu = bounds[0][3]

    for shape in shapes:
        _set_shape_bounds(shape, width_emu=target_w_emu, height_emu=target_h_emu)

    return list(shapes)


def check_bounding_box_collision(b1: Any, b2: Any, tolerance_emu: int = 0) -> bool:
    """Check if two bounding boxes intersect with optional EMU tolerance."""
    l1, t1, w1, h1 = _get_shape_bounds(b1)
    l2, t2, w2, h2 = _get_shape_bounds(b2)

    r1, bot1 = l1 + w1, t1 + h1
    r2, bot2 = l2 + w2, t2 + h2

    overlap_x = min(r1, r2) - max(l1, l2)
    overlap_y = min(bot1, bot2) - max(t1, t2)

    return (overlap_x > tolerance_emu) and (overlap_y > tolerance_emu)


def calculate_overlap_box(b1: Any, b2: Any) -> Optional[BoundingBox]:
    """Compute the intersection bounding box between two shapes in EMU coordinates."""
    l1, t1, w1, h1 = _get_shape_bounds(b1)
    l2, t2, w2, h2 = _get_shape_bounds(b2)

    ox_left = max(l1, l2)
    ox_top = max(t1, t2)
    ox_right = min(l1 + w1, l2 + w2)
    ox_bottom = min(t1 + h1, t2 + h2)

    if ox_right > ox_left and ox_bottom > ox_top:
        return BoundingBox(
            left_emu=ox_left,
            top_emu=ox_top,
            width_emu=ox_right - ox_left,
            height_emu=ox_bottom - ox_top,
        )
    return None


def calculate_overlap_area(b1: Any, b2: Any) -> int:
    """Calculate overlap area between two bounding boxes in EMU squared."""
    box = calculate_overlap_box(b1, b2)
    if box is not None:
        return box.width_emu * box.height_emu
    return 0


def calculate_overlap_area_sq_inches(b1: Any, b2: Any) -> float:
    """Calculate overlap area between two bounding boxes in square inches."""
    area_emu_sq = calculate_overlap_area(b1, b2)
    return float(area_emu_sq) / float(EMU_PER_INCH * EMU_PER_INCH)


def detect_slide_overlaps(
    slide_or_model: Any,
    min_overlap_area_sq_in: float = 0.01,
) -> List[Dict[str, Any]]:
    """Scan all shapes on a slide to detect geometric overlaps exceeding a threshold area.

    Args:
        slide_or_model: SlideModel, python-pptx Slide, or Sequence of shapes.
        min_overlap_area_sq_in: Minimum overlapping square inches to report (default 0.01 sq in).

    Returns:
        List of detected overlap records with shape IDs, names, overlap area, and bounding box.
    """
    if hasattr(slide_or_model, "shapes"):
        shapes = list(slide_or_model.shapes)
    elif isinstance(slide_or_model, (list, tuple)):
        shapes = list(slide_or_model)
    else:
        shapes = []

    overlaps: List[Dict[str, Any]] = []
    n = len(shapes)
    min_area_emu_sq = int(round(min_overlap_area_sq_in * EMU_PER_INCH * EMU_PER_INCH))

    for i in range(n):
        for j in range(i + 1, n):
            s1 = shapes[i]
            s2 = shapes[j]

            # Skip comparing a shape against itself or invalid bounds
            try:
                area_emu_sq = calculate_overlap_area(s1, s2)
            except Exception:
                continue

            if area_emu_sq >= min_area_emu_sq:
                overlap_box = calculate_overlap_box(s1, s2)
                area_sq_in = round(float(area_emu_sq) / float(EMU_PER_INCH * EMU_PER_INCH), 4)
                overlaps.append({
                    "shape_1_id": _get_shape_id(s1),
                    "shape_1_name": _get_shape_name(s1),
                    "shape_2_id": _get_shape_id(s2),
                    "shape_2_name": _get_shape_name(s2),
                    "overlap_area_sq_in": area_sq_in,
                    "overlap_area_emu_sq": area_emu_sq,
                    "overlap_bbox": overlap_box.to_dict() if overlap_box else None,
                })

    return overlaps


def detect_off_slide_shapes(
    slide_or_model: Any,
    slide_width_inches: Optional[float] = None,
    slide_height_inches: Optional[float] = None,
    tolerance_inches: float = 0.01,
) -> List[Dict[str, Any]]:
    """Detect shapes extending beyond slide canvas boundaries.

    Args:
        slide_or_model: SlideModel, python-pptx Slide, or Sequence of shapes.
        slide_width_inches: Slide width in inches (defaults to model width or 10.0 in).
        slide_height_inches: Slide height in inches (defaults to model height or 5.625 in).
        tolerance_inches: Allowed boundary breach tolerance in inches (default 0.01 in).

    Returns:
        List of breach records detailing off-slide edges and distances in inches.
    """
    if slide_width_inches is None:
        slide_width_inches = getattr(slide_or_model, "width_inches", 10.0)
    if slide_height_inches is None:
        slide_height_inches = getattr(slide_or_model, "height_inches", 5.625)

    slide_w_emu = inches_to_emu(slide_width_inches)
    slide_h_emu = inches_to_emu(slide_height_inches)
    tolerance_emu = inches_to_emu(tolerance_inches)

    if hasattr(slide_or_model, "shapes"):
        shapes = list(slide_or_model.shapes)
    elif isinstance(slide_or_model, (list, tuple)):
        shapes = list(slide_or_model)
    else:
        shapes = []

    breach_records: List[Dict[str, Any]] = []

    for shape in shapes:
        try:
            l, t, w, h = _get_shape_bounds(shape)
        except Exception:
            continue

        r = l + w
        b = t + h

        breaches: Dict[str, float] = {}

        if l < -tolerance_emu:
            breaches["left"] = emu_to_inches(abs(l))
        if t < -tolerance_emu:
            breaches["top"] = emu_to_inches(abs(t))
        if r > slide_w_emu + tolerance_emu:
            breaches["right"] = emu_to_inches(r - slide_w_emu)
        if b > slide_h_emu + tolerance_emu:
            breaches["bottom"] = emu_to_inches(b - slide_h_emu)

        if breaches:
            breach_records.append({
                "shape_id": _get_shape_id(shape),
                "shape_name": _get_shape_name(shape),
                "breaches": breaches,
                "bbox": BoundingBox(left_emu=l, top_emu=t, width_emu=w, height_emu=h).to_dict(),
            })

    return breach_records
