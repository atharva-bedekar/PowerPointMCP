"""Direct OpenXML DrawingML manipulation helpers for transparency, gradients, drop shadows, and safe XML modifications."""

import copy
from typing import Any, Callable, Dict, List, Optional, Union
from lxml import etree

# OpenXML XML Namespaces
NAMESPACES = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

A_NS = f"{{{NAMESPACES['a']}}}"
P_NS = f"{{{NAMESPACES['p']}}}"


def _get_element(shape_or_element: Any) -> etree._Element:
    """Extract underlying lxml element from python-pptx Shape or return element directly."""
    if hasattr(shape_or_element, "_element"):
        return shape_or_element._element
    if isinstance(shape_or_element, etree._Element):
        return shape_or_element
    raise TypeError(f"Expected shape or lxml element, got {type(shape_or_element).__name__}")


def get_or_create_spPr(shape_or_element: Any) -> etree._Element:
    """Locate or create the `<p:spPr>` (shape properties) element on a shape node."""
    elem = _get_element(shape_or_element)
    spPr = elem.find(f"{P_NS}spPr")
    if spPr is None:
        spPr = elem.find(".//p:spPr", namespaces=NAMESPACES)
    if spPr is None:
        spPr = etree.SubElement(elem, f"{P_NS}spPr")
    return spPr


def _clean_hex(hex_str: Optional[str], default: str = "000000") -> str:
    """Normalize hex color string to 6-character uppercase hex."""
    if not hex_str:
        return default
    cleaned = str(hex_str).strip().lstrip("#").upper()
    if len(cleaned) == 3:
        cleaned = "".join(c * 2 for c in cleaned)
    if len(cleaned) != 6:
        return default
    return cleaned


def set_shape_transparency(shape_or_element: Any, transparency_percent: float) -> None:
    """Set fill transparency percentage on a shape (0.0 = fully opaque, 100.0 = fully transparent).

    Injects or modifies DrawingML `<a:alpha val="...">` on the solidFill color node.
    `val` is in 1/1000th of a percent (e.g. 50% opacity = 50000).

    Args:
        shape_or_element: python-pptx Shape object or lxml element.
        transparency_percent: Transparency percentage between 0.0 and 100.0.
    """
    spPr = get_or_create_spPr(shape_or_element)
    transparency = max(0.0, min(100.0, float(transparency_percent)))
    opacity_percent = 100.0 - transparency
    alpha_val = int(round(opacity_percent * 1000.0))

    # Look for existing solidFill
    solid_fill = spPr.find(f"{A_NS}solidFill")
    if solid_fill is None:
        # Create solidFill if not present
        solid_fill = etree.SubElement(spPr, f"{A_NS}solidFill")
        color_node = etree.SubElement(solid_fill, f"{A_NS}srgbClr", {"val": "000000"})
    else:
        # Find color element inside solidFill (srgbClr, schemeClr, prstClr, sysClr)
        color_node = None
        for child in solid_fill:
            if child.tag.endswith(("srgbClr", "schemeClr", "prstClr", "sysClr")):
                color_node = child
                break
        if color_node is None:
            color_node = etree.SubElement(solid_fill, f"{A_NS}srgbClr", {"val": "000000"})

    # Find or set alpha child
    alpha_node = color_node.find(f"{A_NS}alpha")
    if transparency == 0.0:
        if alpha_node is not None:
            color_node.remove(alpha_node)
    else:
        if alpha_node is None:
            alpha_node = etree.SubElement(color_node, f"{A_NS}alpha")
        alpha_node.set("val", str(alpha_val))


def set_gradient_fill(
    shape_or_element: Any,
    stops: Optional[List[Dict[str, Any]]] = None,
    angle_deg: float = 90.0,
    start_hex: Optional[str] = None,
    end_hex: Optional[str] = None,
) -> None:
    """Apply a smooth gradient fill to a shape via DrawingML `<a:gradFill>`.

    Supports either a list of stop dictionaries or direct start_hex and end_hex values.

    Args:
        shape_or_element: python-pptx Shape object or lxml element.
        stops: List of stop dicts, e.g. [{"position": 0.0, "color": "FF0000"}, {"position": 1.0, "color": "0000FF"}].
               Position can be 0.0..1.0 or 0..100.
        angle_deg: Linear gradient angle in degrees (default 90.0).
        start_hex: Hex color for gradient start if stops is omitted.
        end_hex: Hex color for gradient end if stops is omitted.
    """
    spPr = get_or_create_spPr(shape_or_element)

    # Remove existing fill types
    fill_tags = [f"{A_NS}solidFill", f"{A_NS}gradFill", f"{A_NS}noFill", f"{A_NS}blipFill", f"{A_NS}pattFill", f"{A_NS}grpFill"]
    for child in list(spPr):
        if child.tag in fill_tags:
            spPr.remove(child)

    grad_fill = etree.SubElement(spPr, f"{A_NS}gradFill")
    gs_lst = etree.SubElement(grad_fill, f"{A_NS}gsLst")

    stop_list: List[Dict[str, Any]] = []
    if stops and len(stops) >= 2:
        stop_list = stops
    else:
        s_hex = _clean_hex(start_hex, "FFFFFF")
        e_hex = _clean_hex(end_hex, "000000")
        stop_list = [
            {"position": 0.0, "color": s_hex},
            {"position": 1.0, "color": e_hex},
        ]

    for stop in stop_list:
        raw_pos = stop.get("position", stop.get("pos", 0.0))
        if isinstance(raw_pos, (int, float)):
            if raw_pos <= 1.0:
                pos_val = int(round(raw_pos * 100000.0))
            elif raw_pos <= 100.0:
                pos_val = int(round(raw_pos * 1000.0))
            else:
                pos_val = int(round(raw_pos))
        else:
            pos_val = 0
        pos_val = max(0, min(100000, pos_val))

        clr_hex = _clean_hex(stop.get("color", stop.get("color_hex", "000000")))

        gs = etree.SubElement(gs_lst, f"{A_NS}gs", {"pos": str(pos_val)})
        srgb_clr = etree.SubElement(gs, f"{A_NS}srgbClr", {"val": clr_hex})

        # Check for stop transparency / alpha if specified
        transp = stop.get("transparency", stop.get("transparency_percent"))
        if transp is not None:
            op_pct = 100.0 - float(transp)
            a_val = int(round(op_pct * 1000.0))
            etree.SubElement(srgb_clr, f"{A_NS}alpha", {"val": str(a_val)})

    # Linear gradient direction: ang in 60000ths of a degree
    ang_val = int(round(angle_deg * 60000.0))
    etree.SubElement(grad_fill, f"{A_NS}lin", {"ang": str(ang_val), "scaled": "0"})


def set_shape_gradient_fill(
    shape_or_element: Any,
    start_hex: str,
    end_hex: str,
    angle_deg: float = 90.0,
) -> None:
    """Convenience helper for two-stop linear gradient fill."""
    set_gradient_fill(shape_or_element, start_hex=start_hex, end_hex=end_hex, angle_deg=angle_deg)


def set_drop_shadow(
    shape_or_element: Any,
    blur_rad_pt: float = 4.0,
    dist_pt: float = 3.0,
    dir_deg: float = 45.0,
    color_hex: str = "000000",
    alpha_percent: float = 40.0,
) -> None:
    """Configure an outer drop shadow effect in `<a:effectLst>/<a:outerShdw>`.

    Args:
        shape_or_element: python-pptx Shape object or lxml element.
        blur_rad_pt: Blur radius in points (default 4.0 pt).
        dist_pt: Shadow distance offset in points (default 3.0 pt).
        dir_deg: Direction angle in degrees (default 45.0 deg).
        color_hex: Shadow RGB hex color (default '000000').
        alpha_percent: Shadow opacity percentage (0 = invisible, 100 = full solid, default 40%).
    """
    spPr = get_or_create_spPr(shape_or_element)

    effect_lst = spPr.find(f"{A_NS}effectLst")
    if effect_lst is None:
        effect_lst = etree.SubElement(spPr, f"{A_NS}effectLst")

    # Remove existing outerShdw if present
    outer_shdw = effect_lst.find(f"{A_NS}outerShdw")
    if outer_shdw is not None:
        effect_lst.remove(outer_shdw)

    blur_rad_emu = int(round(blur_rad_pt * 12700.0))
    dist_emu = int(round(dist_pt * 12700.0))
    dir_val = int(round(dir_deg * 60000.0))

    outer_shdw = etree.SubElement(
        effect_lst,
        f"{A_NS}outerShdw",
        {
            "blurRad": str(blur_rad_emu),
            "dist": str(dist_emu),
            "dir": str(dir_val),
            "algn": "tl",
            "rotWithShape": "0",
        },
    )

    clean_color = _clean_hex(color_hex, "000000")
    srgb_clr = etree.SubElement(outer_shdw, f"{A_NS}srgbClr", {"val": clean_color})

    # Opacity in 1/1000th of percent
    opacity = max(0.0, min(100.0, float(alpha_percent)))
    alpha_val = int(round(opacity * 1000.0))
    etree.SubElement(srgb_clr, f"{A_NS}alpha", {"val": str(alpha_val)})


def set_shape_shadow_effect(
    shape_or_element: Any,
    blur_rad_pt: float = 4.0,
    dist_pt: float = 3.0,
    dir_deg: float = 45.0,
    color_hex: str = "000000",
    alpha_percent: float = 40.0,
) -> None:
    """Alias for set_drop_shadow."""
    set_drop_shadow(shape_or_element, blur_rad_pt, dist_pt, dir_deg, color_hex, alpha_percent)


def get_raw_shape_xml(shape_or_element: Any, pretty_print: bool = True) -> str:
    """Retrieve the raw OpenXML string for a shape or element."""
    elem = _get_element(shape_or_element)
    return etree.tostring(elem, encoding="unicode", pretty_print=pretty_print)


def safe_modify_xml(
    shape_or_element: Any,
    modifier_fn: Callable[[etree._Element], None],
) -> bool:
    """Execute a custom XML mutation function with automatic rollback on error.

    Clones the target node beforehand. If modifier_fn raises an exception or corrupts
    the element structure, the node is restored to its original state.

    Args:
        shape_or_element: python-pptx Shape object or lxml element.
        modifier_fn: Callable taking the lxml element to mutate.

    Returns:
        True if the modification succeeded.

    Raises:
        ValueError if the modification failed and was rolled back.
    """
    elem = _get_element(shape_or_element)
    backup = copy.deepcopy(elem)
    parent = elem.getparent()

    try:
        modifier_fn(elem)
        # Verify valid XML generation
        _ = etree.tostring(elem)
        return True
    except Exception as exc:
        if parent is not None:
            try:
                idx = parent.index(elem)
                parent.remove(elem)
                parent.insert(idx, backup)
            except Exception:
                pass
        if hasattr(shape_or_element, "_element"):
            shape_or_element._element = backup
        raise ValueError(f"XML modification failed: {exc}") from exc
