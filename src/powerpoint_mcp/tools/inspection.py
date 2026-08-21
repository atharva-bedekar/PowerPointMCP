"""Inspection tools for PowerPoint MCP server."""

from functools import wraps
import os
from pathlib import Path
import traceback
from typing import Any, Dict, List, Optional, Union

from pptx import Presentation

from powerpoint_mcp.models.shape import ShapeModel, emu_to_inches
from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.models.presentation import PresentationModel
from powerpoint_mcp.pptx.inspector import (
    inspect_presentation,
    inspect_shape,
    inspect_slide,
    match_shapes,
)
from powerpoint_mcp.rendering.visual_compare import compare_slides
from powerpoint_mcp.tools.versioning import get_session_manager
from powerpoint_mcp.utils.validation import validate_slide


def handle_tool_errors(func):
    """Decorator to catch exceptions and return structured JSON error payloads."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as exc:
            return {
                "success": False,
                "error_type": "FileNotFound",
                "message": str(exc),
                "details": {"path": str(kwargs.get("presentation_path", ""))},
            }
        except IndexError as exc:
            return {
                "success": False,
                "error_type": "SlideNotFound",
                "message": str(exc),
                "details": {"error": str(exc)},
            }
        except ValueError as exc:
            msg = str(exc)
            err_type = "ValidationError"
            if "Shape with ID" in msg or "Shape ID" in msg:
                err_type = "ShapeNotFound"
            elif "Slide number" in msg or "Slide" in msg and "out of range" in msg:
                err_type = "SlideNotFound"
            elif "Session not found" in msg or "active" in msg.lower():
                err_type = "SessionNotFound"
            elif "coordinate" in msg.lower() or "dimension" in msg.lower():
                err_type = "InvalidCoordinate"
            elif "text frame" in msg.lower():
                err_type = "InvalidTextFrame"
            elif "xml" in msg.lower():
                err_type = "OOXMLValidationError"
            return {
                "success": False,
                "error_type": err_type,
                "message": msg,
                "details": {"error": msg},
            }
        except Exception as exc:
            return {
                "success": False,
                "error_type": type(exc).__name__,
                "message": str(exc),
                "details": {"traceback": traceback.format_exc()},
            }
    return wrapper


def _resolve_presentation_path(presentation_path: Optional[str] = None) -> str:
    """Resolve presentation path from argument or active session."""
    if presentation_path:
        p = Path(presentation_path).resolve()
        if not p.exists():
            raise FileNotFoundError(f"Presentation file not found: {p}")
        return str(p)

    mgr = get_session_manager()
    session = mgr.get_current_session()
    if session and session.working_path and Path(session.working_path).exists():
        return str(Path(session.working_path).resolve())

    raise ValueError("No presentation path provided and no active editing session found. Please call ppt_open first.")


@handle_tool_errors
def ppt_inspect_presentation(
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect presentation metadata, slide count, dimensions, layouts, and slide summaries.

    Args:
        presentation_path: Path to presentation. If omitted, uses the active session's working copy.

    Returns:
        Structured dictionary detailing presentation dimensions, slide count, layouts, and titles.
    """
    target_path = _resolve_presentation_path(presentation_path)
    model: PresentationModel = inspect_presentation(target_path)

    return {
        "success": True,
        "path": model.path,
        "slide_count": model.slide_count,
        "width_inches": model.width_inches,
        "height_inches": model.height_inches,
        "slide_width_emu": model.width_emu,
        "slide_height_emu": model.height_emu,
        "layouts": model.layouts,
        "theme": model.theme_name or "Default",
        "metadata": model.metadata.to_dict() if model.metadata else {},
        "titles": [s.title for s in model.slides if s.title],
        "slides": model.slide_titles if model.slide_titles else [
            {
                "slide_number": s.slide_number,
                "slide_id": s.slide_id,
                "title": s.title,
                "layout_name": s.layout_name,
                "shape_count": s.shape_count,
            }
            for s in model.slides
        ],
    }


@handle_tool_errors
def ppt_inspect_slide(
    slide_number: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Inspect all shapes on a specific slide with coordinates, semantic roles, typography, and styling.

    Args:
        slide_number: 1-indexed slide number.
        presentation_path: Path to presentation. If omitted, uses active session.

    Returns:
        Structured dictionary containing slide layout, title, dimensions, and shape collection.
    """
    if slide_number < 1:
        raise IndexError(f"Slide number must be >= 1, got {slide_number}")

    target_path = _resolve_presentation_path(presentation_path)
    slide_model: SlideModel = inspect_slide(target_path, slide_number)

    return {
        "success": True,
        "slide_number": slide_model.slide_number,
        "slide_id": slide_model.slide_id,
        "layout_name": slide_model.layout_name,
        "title": slide_model.title,
        "shape_count": slide_model.shape_count,
        "width_inches": slide_model.width_inches,
        "height_inches": slide_model.height_inches,
        "width_emu": slide_model.width_emu,
        "height_emu": slide_model.height_emu,
        "has_notes": slide_model.has_notes,
        "notes": slide_model.notes,
        "shapes": [s.to_dict() for s in slide_model.shapes],
    }


@handle_tool_errors
def ppt_inspect_shape(
    slide_number: int,
    shape_id: int,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Get deep inspection details for a single shape on a slide.

    Args:
        slide_number: 1-indexed slide number.
        shape_id: ID of the shape to inspect.
        presentation_path: Path to presentation. If omitted, uses active session.

    Returns:
        Complete shape details with bounding box, text frames, paragraphs, runs, fills, and lines.
    """
    if slide_number < 1:
        raise IndexError(f"Slide number must be >= 1, got {slide_number}")

    target_path = _resolve_presentation_path(presentation_path)
    shape_model: ShapeModel = inspect_shape(target_path, slide_number, shape_id)

    return {
        "success": True,
        "slide_number": slide_number,
        "shape_id": shape_id,
        "shape": shape_model.to_dict(),
    }


@handle_tool_errors
def ppt_compare_slides(
    slide_a: int,
    slide_b: int,
    match_shapes_flag: bool = True,
    render_diff: bool = False,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Compare geometric, typographic, and semantic layout properties between two slides.

    Args:
        slide_a: 1-indexed reference slide number.
        slide_b: 1-indexed comparison slide number.
        match_shapes_flag: Whether to perform multi-factor semantic shape matching.
        render_diff: Whether to perform pixel visual diffing if renderers are available.
        presentation_path: Path to presentation. If omitted, uses active session.

    Returns:
        Slide comparison report detailing matched shapes, layout differences, and similarity score.
    """
    if slide_a < 1 or slide_b < 1:
        raise IndexError(f"Slide numbers must be >= 1, got slide_a={slide_a}, slide_b={slide_b}")

    target_path = _resolve_presentation_path(presentation_path)
    slide_a_model = inspect_slide(target_path, slide_a)
    slide_b_model = inspect_slide(target_path, slide_b)

    comp_result = compare_slides(slide_a_model, slide_b_model)
    res_dict = comp_result.to_dict()

    return {
        "success": True,
        **res_dict,
    }


@handle_tool_errors
def ppt_validate_slide(
    slide_number: int,
    rules: Optional[List[str]] = None,
    presentation_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Run rule-based geometric and typographic validation on a slide.

    Detects overlaps (VAL-01), boundary clipping (VAL-02), off-slide objects (VAL-03),
    text overflow (VAL-04), tiny fonts (VAL-05), irregular rotation (VAL-08), and duplicate shapes (VAL-07).

    Args:
        slide_number: 1-indexed slide number.
        rules: Optional list of specific rule IDs to check (e.g. ['VAL-01', 'VAL-02']).
        presentation_path: Path to presentation. If omitted, uses active session.

    Returns:
        Slide validation report with is_valid, warning_count, warnings list, and slide metrics.
    """
    if slide_number < 1:
        raise IndexError(f"Slide number must be >= 1, got {slide_number}")

    target_path = _resolve_presentation_path(presentation_path)
    slide_model = inspect_slide(target_path, slide_number)
    val_result = validate_slide(slide_model, rules=rules)

    return {
        "success": True,
        **val_result.to_dict(),
    }
