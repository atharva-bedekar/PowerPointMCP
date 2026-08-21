"""Helpers for inspecting embedded OpenXML relationships, pictures, and hyperlinks."""

import hashlib
from typing import Any, Dict, List, Optional
from pptx.enum.shapes import MSO_SHAPE_TYPE


def inspect_slide_relationships(slide: Any) -> List[Dict[str, Any]]:
    """Extract all OpenXML part relationships associated with a slide."""
    relationships: List[Dict[str, Any]] = []
    try:
        part = getattr(slide, "part", None)
        if part and hasattr(part, "rels"):
            for r_id, rel in part.rels.items():
                relationships.append({
                    "r_id": r_id,
                    "rel_type": getattr(rel, "reltype", None),
                    "target_ref": getattr(rel, "target_ref", None),
                    "is_external": getattr(rel, "is_external", False),
                })
    except Exception:
        pass
    return relationships


def extract_embedded_images(slide_or_presentation: Any) -> List[Dict[str, Any]]:
    """Extract embedded images metadata from a slide or an entire presentation."""
    images: List[Dict[str, Any]] = []

    # If it is a presentation
    if hasattr(slide_or_presentation, "slides"):
        for slide in slide_or_presentation.slides:
            images.extend(extract_embedded_images(slide))
        return images

    # If it is a slide
    shapes = getattr(slide_or_presentation, "shapes", [])
    for shape in shapes:
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or getattr(shape, "shape_type", None) == 13:
                if hasattr(shape, "image"):
                    img = shape.image
                    blob = img.blob
                    img_hash = hashlib.sha256(blob).hexdigest()
                    width_px, height_px = img.size if hasattr(img, "size") else (None, None)
                    images.append({
                        "shape_id": shape.shape_id,
                        "shape_name": shape.name,
                        "content_type": getattr(img, "content_type", None),
                        "extension": getattr(img, "ext", None),
                        "sha256": img_hash,
                        "size_bytes": len(blob),
                        "width_px": width_px,
                        "height_px": height_px,
                    })
        except Exception:
            pass

    return images


def extract_hyperlinks(slide: Any) -> List[Dict[str, Any]]:
    """Extract all hyperlinks configured on shapes and text runs of a slide."""
    links: List[Dict[str, Any]] = []
    shapes = getattr(slide, "shapes", [])
    for shape in shapes:
        # Check shape-level click action / hyperlink
        try:
            if hasattr(shape, "click_action") and shape.click_action.hyperlink is not None:
                hl = shape.click_action.hyperlink
                if hl.address:
                    links.append({
                        "shape_id": shape.shape_id,
                        "shape_name": shape.name,
                        "type": "shape",
                        "text": getattr(shape, "text", ""),
                        "address": hl.address,
                    })
        except Exception:
            pass

        # Check text run hyperlinks
        if getattr(shape, "has_text_frame", False):
            try:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if hasattr(run, "hyperlink") and run.hyperlink is not None:
                            if run.hyperlink.address:
                                links.append({
                                    "shape_id": shape.shape_id,
                                    "shape_name": shape.name,
                                    "type": "run",
                                    "text": run.text,
                                    "address": run.hyperlink.address,
                                })
            except Exception:
                pass

    return links


def get_image_part_from_shape(shape: Any) -> Optional[Any]:
    """Retrieve the underlying ImagePart object from a picture shape."""
    try:
        if hasattr(shape, "image") and hasattr(shape.image, "_image_part"):
            return shape.image._image_part
        if hasattr(shape, "image") and hasattr(shape.image, "part"):
            return shape.image.part
    except Exception:
        pass
    return None
