"""PPTX Inspection Engine: metadata, slide trees, shape hierarchies, semantic role inference, and shape matching."""

import difflib
import math
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.presentation import Presentation as PresentationClass

from powerpoint_mcp.models.presentation import (
    PresentationMetadata,
    PresentationModel,
)
from powerpoint_mcp.models.shape import (
    BoundingBox,
    SemanticRole,
    ShapeModel,
    ShapeType,
    emu_to_inches,
    emu_to_pt,
)
from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.pptx.relationships import (
    extract_embedded_images,
    extract_hyperlinks,
    inspect_slide_relationships,
)
from powerpoint_mcp.pptx.styles import (
    extract_fill_style,
    extract_line_style,
    extract_shape_properties,
    extract_text_frame,
)


def _load_presentation(path_or_prs: Union[str, Path, PresentationClass]) -> Tuple[PresentationClass, str]:
    """Ensure a python-pptx Presentation instance is available."""
    if isinstance(path_or_prs, (str, Path)):
        prs_path = str(Path(path_or_prs).resolve())
        prs = Presentation(prs_path)
        return prs, prs_path
    elif hasattr(path_or_prs, "slides"):
        return path_or_prs, getattr(path_or_prs, "_path", "<in-memory>")
    else:
        raise ValueError(f"Invalid presentation input: expected file path or Presentation instance, got {type(path_or_prs)}.")


def map_shape_type(shape: Any) -> ShapeType:
    """Map python-pptx shape to structural ShapeType enum."""
    try:
        shape_type_val = getattr(shape, "shape_type", None)
        if shape_type_val == MSO_SHAPE_TYPE.AUTO_SHAPE or shape_type_val == 1:
            return ShapeType.AUTO_SHAPE
        elif shape_type_val == MSO_SHAPE_TYPE.TEXT_BOX or shape_type_val == 17:
            return ShapeType.TEXT_BOX
        elif shape_type_val == MSO_SHAPE_TYPE.PICTURE or shape_type_val == 13:
            return ShapeType.PICTURE
        elif shape_type_val == MSO_SHAPE_TYPE.GROUP or shape_type_val == 6:
            return ShapeType.GROUP
        elif shape_type_val == MSO_SHAPE_TYPE.TABLE or shape_type_val == 19 or getattr(shape, "has_table", False):
            return ShapeType.TABLE
        elif shape_type_val == MSO_SHAPE_TYPE.CHART or shape_type_val == 3 or getattr(shape, "has_chart", False):
            return ShapeType.CHART
        elif shape_type_val in (MSO_SHAPE_TYPE.LINE, 9, MSO_SHAPE_TYPE.FREEFORM, 5):
            return ShapeType.CONNECTOR
        elif shape_type_val == MSO_SHAPE_TYPE.MEDIA or shape_type_val == 16:
            return ShapeType.MEDIA
        elif shape_type_val == MSO_SHAPE_TYPE.PLACEHOLDER or shape_type_val == 14:
            if getattr(shape, "has_text_frame", False):
                return ShapeType.TEXT_BOX
            elif getattr(shape, "has_table", False):
                return ShapeType.TABLE
            elif getattr(shape, "has_chart", False):
                return ShapeType.CHART
            elif hasattr(shape, "image"):
                return ShapeType.PICTURE
            return ShapeType.AUTO_SHAPE
    except Exception:
        pass
    return ShapeType.UNKNOWN


def infer_semantic_role(shape: Any, slide_width_emu: int, slide_height_emu: int) -> SemanticRole:
    """Infer the conservative semantic role of a shape using a 5-stage rule cascade."""
    # -------------------------------------------------------------------------
    # Stage 1: Placeholder Examination
    # -------------------------------------------------------------------------
    if getattr(shape, "is_placeholder", False):
        try:
            ph_type = shape.placeholder_format.type
            # Title placeholders
            if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, 1, 3):
                return SemanticRole.TITLE
            elif ph_type in (PP_PLACEHOLDER.SUBTITLE, 4):
                return SemanticRole.SUBTITLE
            elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, 2, 7):
                return SemanticRole.BODY
            elif ph_type in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.HEADER, 15, 16, 14, 10):
                return SemanticRole.FOOTER
            elif ph_type in (PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.BITMAP, 18, 9):
                return SemanticRole.IMAGE
            elif ph_type in (PP_PLACEHOLDER.TABLE, 12):
                return SemanticRole.TABLE
            elif ph_type in (PP_PLACEHOLDER.CHART, 8):
                return SemanticRole.CHART
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # Stage 2: Non-Placeholder Structural Types
    # -------------------------------------------------------------------------
    shape_type_val = getattr(shape, "shape_type", None)
    if shape_type_val == MSO_SHAPE_TYPE.PICTURE or shape_type_val == 13 or hasattr(shape, "image"):
        return SemanticRole.IMAGE

    if getattr(shape, "has_table", False) or shape_type_val == MSO_SHAPE_TYPE.TABLE or shape_type_val == 19:
        return SemanticRole.TABLE

    if getattr(shape, "has_chart", False) or shape_type_val == MSO_SHAPE_TYPE.CHART or shape_type_val == 3:
        return SemanticRole.CHART

    if shape_type_val == MSO_SHAPE_TYPE.GROUP or shape_type_val == 6:
        # Group shapes containing connectors or child shapes represent diagrams
        return SemanticRole.DIAGRAM

    if shape_type_val in (MSO_SHAPE_TYPE.LINE, 9, MSO_SHAPE_TYPE.FREEFORM, 5):
        return SemanticRole.DIAGRAM

    # -------------------------------------------------------------------------
    # Stage 3: Spatial & Typographical Heuristics (for Text Shapes)
    # -------------------------------------------------------------------------
    has_text = False
    text_content = ""
    max_font_size = 18.0
    paragraphs_count = 0

    if getattr(shape, "has_text_frame", False):
        try:
            tf = shape.text_frame
            text_content = (tf.text or "").strip()
            if text_content:
                has_text = True
                paragraphs_count = len(tf.paragraphs)
                found_sizes = []
                for p in tf.paragraphs:
                    p_font = getattr(p, "font", None)
                    if p_font and p_font.size is not None:
                        try:
                            found_sizes.append(float(p_font.size.pt))
                        except Exception:
                            found_sizes.append(emu_to_pt(int(p_font.size)))
                    for r in p.runs:
                        r_font = getattr(r, "font", None)
                        if r_font and r_font.size is not None:
                            try:
                                found_sizes.append(float(r_font.size.pt))
                            except Exception:
                                found_sizes.append(emu_to_pt(int(r_font.size)))
                if found_sizes:
                    max_font_size = max(found_sizes)
        except Exception:
            pass

    if has_text:
        slide_h = float(slide_height_emu) if slide_height_emu > 0 else 5143500.0
        norm_top = getattr(shape, "top", 0) / slide_h
        shape_name = getattr(shape, "name", "")

        # Rule 3A: Title Detection
        if (norm_top < 0.22 and max_font_size >= 24) or ("Title" in shape_name and norm_top < 0.35):
            return SemanticRole.TITLE
        if norm_top < 0.16 and max_font_size >= 20:
            return SemanticRole.TITLE

        # Rule 3B: Subtitle Detection
        shape_h_norm = getattr(shape, "height", 0) / slide_h
        if shape_h_norm <= 0.20 and ((0.12 <= norm_top < 0.38 and 14 <= max_font_size < 24) or ("Subtitle" in shape_name and norm_top < 0.45)):
            return SemanticRole.SUBTITLE

        # Rule 3C: Footer Detection
        if norm_top >= 0.85 or any(k in shape_name for k in ("Footer", "Slide Number", "Date", "Page Number")):
            return SemanticRole.FOOTER

        # Rule 3D: Body Content
        if paragraphs_count > 1 or norm_top >= 0.25:
            return SemanticRole.BODY

        if norm_top < 0.25:
            if max_font_size >= 20:
                return SemanticRole.TITLE
            return SemanticRole.BODY

    # -------------------------------------------------------------------------
    # Stage 4: Default Fallback
    # -------------------------------------------------------------------------
    return SemanticRole.UNKNOWN


def inspect_shape(
    path_or_prs: Union[str, Path, PresentationClass],
    slide_number: int,
    shape_id: int,
) -> ShapeModel:
    """Inspect a specific shape by ID on a given 1-indexed slide."""
    prs, _ = _load_presentation(path_or_prs)
    if slide_number < 1 or slide_number > len(prs.slides):
        raise ValueError(f"Slide number {slide_number} is out of bounds (presentation has {len(prs.slides)} slides).")

    slide = prs.slides[slide_number - 1]
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)

    target_shape = None
    target_z = 0
    for idx, shape in enumerate(slide.shapes):
        if shape.shape_id == shape_id:
            target_shape = shape
            target_z = idx
            break

    if target_shape is None:
        raise ValueError(f"Shape ID {shape_id} does not exist on slide {slide_number}.")

    shape = target_shape
    bbox = BoundingBox(
        left_emu=int(getattr(shape, "left", 0)),
        top_emu=int(getattr(shape, "top", 0)),
        width_emu=int(getattr(shape, "width", 0)),
        height_emu=int(getattr(shape, "height", 0)),
    )
    rotation = float(getattr(shape, "rotation", 0.0) or 0.0)
    shape_type = map_shape_type(shape)
    role = infer_semantic_role(shape, slide_w, slide_h)

    text_frame = None
    if getattr(shape, "has_text_frame", False):
        text_frame = extract_text_frame(shape.text_frame)

    fill = extract_fill_style(shape)
    line = extract_line_style(shape)
    props = extract_shape_properties(shape)

    image_metadata = None
    if shape_type == ShapeType.PICTURE or hasattr(shape, "image"):
        try:
            img = shape.image
            image_metadata = {
                "content_type": getattr(img, "content_type", None),
                "extension": getattr(img, "ext", None),
                "size_bytes": len(img.blob),
            }
        except Exception:
            pass

    table_metadata = props.get("table")
    chart_metadata = props.get("chart")

    return ShapeModel(
        shape_id=shape.shape_id,
        name=shape.name or f"Shape {shape.shape_id}",
        shape_type=shape_type,
        semantic_role=role,
        bbox=bbox,
        rotation=rotation,
        z_order=target_z,
        text_frame=text_frame,
        fill=fill,
        line=line,
        properties=props,
        image_metadata=image_metadata,
        table_metadata=table_metadata,
        chart_metadata=chart_metadata,
    )


def inspect_slide(
    path_or_prs: Any,
    slide_number: int,
) -> SlideModel:
    """Inspect an individual 1-indexed slide and all of its shapes."""
    if hasattr(path_or_prs, "shapes") and not hasattr(path_or_prs, "slides"):
        slide = path_or_prs
        slide_w_emu = 12192000
        slide_h_emu = 6858000
        try:
            prs_part = slide.part.package.presentation_part
            slide_w_emu = int(prs_part.presentation.slide_width)
            slide_h_emu = int(prs_part.presentation.slide_height)
        except Exception:
            pass
    else:
        prs, _ = _load_presentation(path_or_prs)
        if slide_number < 1 or slide_number > len(prs.slides):
            raise ValueError(f"Slide number {slide_number} is out of bounds (presentation has {len(prs.slides)} slides).")

        slide = prs.slides[slide_number - 1]
        slide_w_emu = int(prs.slide_width)
        slide_h_emu = int(prs.slide_height)

    slide_w_in = emu_to_inches(slide_w_emu)
    slide_h_in = emu_to_inches(slide_h_emu)

    layout_name = "Custom"
    try:
        if slide.slide_layout and hasattr(slide.slide_layout, "name"):
            layout_name = slide.slide_layout.name or "Custom"
    except Exception:
        pass

    shapes: List[ShapeModel] = []
    slide_title: Optional[str] = None

    for idx, shape in enumerate(slide.shapes):
        bbox = BoundingBox(
            left_emu=int(getattr(shape, "left", 0)),
            top_emu=int(getattr(shape, "top", 0)),
            width_emu=int(getattr(shape, "width", 0)),
            height_emu=int(getattr(shape, "height", 0)),
        )
        rotation = float(getattr(shape, "rotation", 0.0) or 0.0)
        shape_type = map_shape_type(shape)
        role = infer_semantic_role(shape, slide_w_emu, slide_h_emu)

        text_frame = None
        if getattr(shape, "has_text_frame", False):
            text_frame = extract_text_frame(shape.text_frame)
            if role == SemanticRole.TITLE and slide_title is None and text_frame.text.strip():
                slide_title = text_frame.text.strip()

        fill = extract_fill_style(shape)
        line = extract_line_style(shape)
        props = extract_shape_properties(shape)

        image_metadata = None
        if shape_type == ShapeType.PICTURE or hasattr(shape, "image"):
            try:
                img = shape.image
                image_metadata = {
                    "content_type": getattr(img, "content_type", None),
                    "extension": getattr(img, "ext", None),
                    "size_bytes": len(img.blob),
                }
            except Exception:
                pass

        shapes.append(
            ShapeModel(
                shape_id=shape.shape_id,
                name=shape.name or f"Shape {shape.shape_id}",
                shape_type=shape_type,
                semantic_role=role,
                bbox=bbox,
                rotation=rotation,
                z_order=idx,
                text_frame=text_frame,
                fill=fill,
                line=line,
                properties=props,
                image_metadata=image_metadata,
                table_metadata=props.get("table"),
                chart_metadata=props.get("chart"),
            )
        )

    # If title wasn't found from TITLE role, check shape named title or title shape or top-most text shape
    if slide_title is None:
        try:
            if slide.shapes.title and slide.shapes.title.text_frame:
                slide_title = slide.shapes.title.text_frame.text.strip()
        except Exception:
            pass

    if slide_title is None:
        for s in shapes:
            if s.text_frame and s.text_frame.text.strip() and s.bbox.top_inches < 1.5:
                slide_title = s.text_frame.text.strip().split("\n")[0]
                break

    notes_text: Optional[str] = None
    has_notes = False
    try:
        if getattr(slide, "has_notes_slide", False) and slide.notes_slide:
            notes_tf = getattr(slide.notes_slide, "notes_text_frame", None)
            if notes_tf and notes_tf.text.strip():
                notes_text = notes_tf.text.strip()
                has_notes = True
    except Exception:
        pass

    return SlideModel(
        slide_number=slide_number,
        slide_id=slide.slide_id,
        layout_name=layout_name,
        title=slide_title,
        width_inches=slide_w_in,
        height_inches=slide_h_in,
        width_emu=slide_w_emu,
        height_emu=slide_h_emu,
        shapes=shapes,
        notes=notes_text,
        has_notes=has_notes,
        notes_text=notes_text,
    )


def inspect_presentation(
    path_or_prs: Union[str, Path, PresentationClass],
) -> PresentationModel:
    """Inspect presentation metadata, slide titles, dimensions, and full slide trees."""
    prs, prs_path = _load_presentation(path_or_prs)

    slide_w_emu = int(prs.slide_width)
    slide_h_emu = int(prs.slide_height)
    slide_w_in = emu_to_inches(slide_w_emu)
    slide_h_in = emu_to_inches(slide_h_emu)

    layouts: List[str] = []
    try:
        for layout in prs.slide_layouts:
            if layout.name and layout.name not in layouts:
                layouts.append(layout.name)
    except Exception:
        pass

    slides: List[SlideModel] = []
    slide_titles: List[Dict[str, Any]] = []

    for i in range(len(prs.slides)):
        slide_model = inspect_slide(prs, i + 1)
        slides.append(slide_model)
        slide_titles.append({
            "slide_number": i + 1,
            "slide_id": slide_model.slide_id,
            "title": slide_model.title,
            "layout_name": slide_model.layout_name,
            "shape_count": slide_model.shape_count,
        })

    # Core properties / metadata
    metadata = PresentationMetadata()
    try:
        cp = prs.core_properties
        metadata.title = getattr(cp, "title", None)
        metadata.author = getattr(cp, "author", None)
        metadata.subject = getattr(cp, "subject", None)
        metadata.category = getattr(cp, "category", None)
        metadata.comments = getattr(cp, "comments", None)
        metadata.revision = getattr(cp, "revision", None)
        if getattr(cp, "created", None):
            metadata.created = str(cp.created)
        if getattr(cp, "modified", None):
            metadata.modified = str(cp.modified)
    except Exception:
        pass

    return PresentationModel(
        path=prs_path,
        width_inches=slide_w_in,
        height_inches=slide_h_in,
        width_emu=slide_w_emu,
        height_emu=slide_h_emu,
        slide_count=len(slides),
        theme_name=None,
        layouts=layouts,
        slides=slides,
        slide_titles=slide_titles,
        metadata=metadata,
    )


def match_shapes(
    slide_a: SlideModel,
    slide_b: SlideModel,
    min_confidence: float = 0.40,
) -> List[Dict[str, Any]]:
    """Match shapes between two slides using a multi-factor scoring heuristic.

    Factors:
    - Semantic Role (weight 0.25)
    - Text Similarity (weight 0.25)
    - Relative Position (weight 0.20)
    - Shape Type (weight 0.15)
    - Relative Dimensions (weight 0.10)
    - Shape Name Similarity (weight 0.05)
    """
    weights = {
        "role": 0.25,
        "text": 0.25,
        "position": 0.20,
        "type": 0.15,
        "dimensions": 0.10,
        "name": 0.05,
    }

    slide_w = float(max(slide_a.width_emu, slide_b.width_emu, 1))
    slide_h = float(max(slide_a.height_emu, slide_b.height_emu, 1))

    candidates: List[Dict[str, Any]] = []

    for sa in slide_a.shapes:
        for sb in slide_b.shapes:
            # 1. Semantic Role Score
            if sa.semantic_role == sb.semantic_role and sa.semantic_role != SemanticRole.UNKNOWN:
                role_score = 1.0
            elif sa.semantic_role == SemanticRole.UNKNOWN and sb.semantic_role == SemanticRole.UNKNOWN:
                role_score = 0.20
            elif sa.semantic_role == SemanticRole.UNKNOWN or sb.semantic_role == SemanticRole.UNKNOWN:
                role_score = 0.05
            else:
                role_score = 0.0

            # 2. Text Similarity Score
            text_a = (sa.text_frame.text if sa.text_frame else "").strip().lower()
            text_b = (sb.text_frame.text if sb.text_frame else "").strip().lower()
            if not text_a and not text_b:
                text_score = 1.0  # Both have no text (e.g. pictures/shapes matching each other)
            elif not text_a or not text_b:
                text_score = 0.0
            else:
                text_score = difflib.SequenceMatcher(None, text_a, text_b).ratio()

            # 3. Relative Position Score
            center_dx = (sa.bbox.center_x_emu - sb.bbox.center_x_emu) / slide_w
            center_dy = (sa.bbox.center_y_emu - sb.bbox.center_y_emu) / slide_h
            center_dist = math.sqrt(center_dx * center_dx + center_dy * center_dy)
            pos_score = max(0.0, 1.0 - (2.0 * center_dist))

            # 4. Shape Type Score
            type_score = 1.0 if sa.shape_type == sb.shape_type else 0.0

            # 5. Relative Dimensions Score
            dim_dw = abs(sa.bbox.width_emu - sb.bbox.width_emu) / slide_w
            dim_dh = abs(sa.bbox.height_emu - sb.bbox.height_emu) / slide_h
            dim_score = max(0.0, 1.0 - (dim_dw + dim_dh))

            # 6. Shape Name Similarity Score
            name_a = sa.name.lower()
            name_b = sb.name.lower()
            name_score = difflib.SequenceMatcher(None, name_a, name_b).ratio()

            total_score = (
                weights["role"] * role_score
                + weights["text"] * text_score
                + weights["position"] * pos_score
                + weights["type"] * type_score
                + weights["dimensions"] * dim_score
                + weights["name"] * name_score
            )

            reasoning_parts = []
            if role_score == 1.0:
                reasoning_parts.append(f"identical role '{sa.semantic_role.value}'")
            if text_score > 0.8:
                reasoning_parts.append(f"high text similarity ({text_score:.2f})")
            if pos_score > 0.8:
                reasoning_parts.append("closely aligned position")
            if type_score == 1.0:
                reasoning_parts.append(f"matching type '{sa.shape_type.value}'")

            reasoning = "; ".join(reasoning_parts) if reasoning_parts else "partial geometric and structural similarity"

            candidates.append({
                "shape_a_id": sa.shape_id,
                "shape_b_id": sb.shape_id,
                "shape_a_name": sa.name,
                "shape_b_name": sb.name,
                "confidence_score": round(total_score, 4),
                "factors": {
                    "role_score": round(role_score, 4),
                    "text_score": round(text_score, 4),
                    "position_score": round(pos_score, 4),
                    "type_score": round(type_score, 4),
                    "dimensions_score": round(dim_score, 4),
                    "name_score": round(name_score, 4),
                },
                "reasoning": reasoning,
            })

    # Greedy bipartite assignment
    candidates.sort(key=lambda c: c["confidence_score"], reverse=True)
    matched_a: Set[int] = set()
    matched_b: Set[int] = set()
    final_matches: List[Dict[str, Any]] = []

    for cand in candidates:
        if cand["confidence_score"] < min_confidence:
            break
        a_id = cand["shape_a_id"]
        b_id = cand["shape_b_id"]
        if a_id not in matched_a and b_id not in matched_b:
            matched_a.add(a_id)
            matched_b.add(b_id)
            final_matches.append(cand)

    return final_matches


class PPTXInspector:
    """Convenience class interface for PowerPoint inspection operations."""

    @staticmethod
    def inspect_presentation(path_or_prs: Union[str, Path, PresentationClass]) -> PresentationModel:
        """Inspect entire presentation."""
        return inspect_presentation(path_or_prs)

    @staticmethod
    def inspect_slide(path_or_prs: Union[str, Path, PresentationClass], slide_number: int) -> SlideModel:
        """Inspect a single slide (1-indexed)."""
        return inspect_slide(path_or_prs, slide_number)

    @staticmethod
    def inspect_shape(path_or_prs: Union[str, Path, PresentationClass], slide_number: int, shape_id: int) -> ShapeModel:
        """Inspect a specific shape by ID on a slide."""
        return inspect_shape(path_or_prs, slide_number, shape_id)

    @staticmethod
    def match_shapes(slide_a: SlideModel, slide_b: SlideModel, min_confidence: float = 0.40) -> List[Dict[str, Any]]:
        """Match shapes across slides."""
        return match_shapes(slide_a, slide_b, min_confidence=min_confidence)
