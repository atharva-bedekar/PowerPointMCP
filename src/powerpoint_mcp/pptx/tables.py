"""Core operations for PowerPoint tables: inspection, cell mutation, geometry, styling, and cell merging."""

from typing import Any, Dict, List, Optional, Tuple, Union
import re

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

from powerpoint_mcp.models.shape import (
    emu_to_inches,
    emu_to_pt,
    inches_to_emu,
    pt_to_emu,
)


def _hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert hex string (e.g. '#1E3A8A' or '1E3A8A') to pptx RGBColor."""
    clean = hex_str.strip().lstrip("#")
    if len(clean) == 3:
        clean = "".join(c * 2 for c in clean)
    if len(clean) != 6:
        raise ValueError(f"Invalid hex color: {hex_str}")
    r = int(clean[0:2], 16)
    g = int(clean[2:4], 16)
    b = int(clean[4:6], 16)
    return RGBColor(r, g, b)


def _rgb_to_hex(rgb: Any) -> Optional[str]:
    """Convert pptx RGBColor to hex string."""
    if rgb is None:
        return None
    try:
        return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    except Exception:
        return None


def _find_table_shape(slide: Any, table_shape_id: Optional[int] = None) -> Any:
    """Locate table shape on slide by ID or return the first table found."""
    for shape in slide.shapes:
        has_tbl = getattr(shape, "has_table", False) or getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.TABLE
        if has_tbl:
            if table_shape_id is None or shape.shape_id == table_shape_id:
                return shape

    if table_shape_id is not None:
        raise ValueError(f"Table with shape ID {table_shape_id} not found on slide")
    raise ValueError("No table found on slide")


def inspect_table_cells(
    slide: Any,
    table_shape_id: Optional[int] = None,
    detail: str = "compact",
) -> Dict[str, Any]:
    """Inspect a PowerPoint table at cell level.

    Args:
        slide: Target pptx Slide.
        table_shape_id: Optional specific shape ID. If omitted, finds first table on slide.
        detail: 'compact' for agent-friendly summary, or 'full' for deep diagnostics.

    Returns:
        Structured dictionary with table dimensions, geometry, and cell data.
    """
    shape = _find_table_shape(slide, table_shape_id)
    table = shape.table

    num_rows = len(table.rows)
    num_cols = len(table.columns)

    col_widths = [emu_to_inches(c.width) for c in table.columns]
    row_heights = [emu_to_inches(r.height) for r in table.rows]

    bbox = {
        "x": emu_to_inches(shape.left),
        "y": emu_to_inches(shape.top),
        "width": emu_to_inches(shape.width),
        "height": emu_to_inches(shape.height),
    }

    compact_lines = []
    cells_data = []

    for r_idx, row in enumerate(table.rows):
        row_compact_entries = []
        for c_idx, col in enumerate(table.columns):
            cell = table.cell(r_idx, c_idx)
            text_val = cell.text.strip()
            # Clean up display text for compact view
            display_txt = text_val.replace("\n", " ").strip()
            if len(display_txt) > 40:
                display_txt = display_txt[:37] + "..."
            row_compact_entries.append(f'C{c_idx + 1}: "{display_txt}"')

            if detail == "full":
                font_name = None
                font_size = None
                bold = None
                italic = None
                font_color = None
                align_str = None

                if cell.text_frame and cell.text_frame.paragraphs:
                    first_p = cell.text_frame.paragraphs[0]
                    if first_p.alignment is not None:
                        align_str = str(first_p.alignment).split(" ")[0].lower()
                    if first_p.runs:
                        r0 = first_p.runs[0]
                        font_name = r0.font.name
                        if r0.font.size is not None:
                            font_size = emu_to_pt(r0.font.size)
                        bold = r0.font.bold
                        italic = r0.font.italic
                        if r0.font.color and getattr(r0.font.color, "rgb", None):
                            font_color = _rgb_to_hex(r0.font.color.rgb)

                fill_hex = None
                try:
                    if cell.fill and getattr(cell.fill, "type", None):
                        if getattr(cell.fill, "fore_color", None) and getattr(cell.fill.fore_color, "rgb", None):
                            fill_hex = _rgb_to_hex(cell.fill.fore_color.rgb)
                except Exception:
                    pass

                cells_data.append({
                    "row": r_idx,
                    "column": c_idx,
                    "text": text_val,
                    "font_name": font_name,
                    "font_size": font_size,
                    "bold": bold,
                    "italic": italic,
                    "font_color": font_color,
                    "alignment": align_str,
                    "fill": fill_hex,
                    "is_merge_origin": getattr(cell, "is_merge_origin", False),
                    "is_spanned": getattr(cell, "is_spanned", False),
                    "margin_left": emu_to_inches(cell.margin_left) if cell.margin_left else None,
                    "margin_right": emu_to_inches(cell.margin_right) if cell.margin_right else None,
                    "margin_top": emu_to_inches(cell.margin_top) if cell.margin_top else None,
                    "margin_bottom": emu_to_inches(cell.margin_bottom) if cell.margin_bottom else None,
                })

        compact_lines.append(f"  R{r_idx + 1}: " + ", ".join(row_compact_entries))

    summary = (
        f"Table Shape ID: {shape.shape_id}\n"
        f"Grid: {num_rows} rows x {num_cols} columns\n"
        f"Bounding box: left={bbox['x']}\", top={bbox['y']}\", width={bbox['width']}\", height={bbox['height']}\"\n"
        f"Column widths: {col_widths}\n"
        f"Row heights: {row_heights}\n"
        f"Content:\n" + "\n".join(compact_lines)
    )

    result = {
        "shape_id": shape.shape_id,
        "name": shape.name,
        "rows": num_rows,
        "columns": num_cols,
        "bbox": bbox,
        "column_widths": col_widths,
        "row_heights": row_heights,
        "summary": summary,
    }
    if detail == "full":
        result["cells"] = cells_data

    return result


def batch_modify_table_cells(
    slide: Any,
    table_shape_id: int,
    mutations: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """Batch modify multiple table cells atomically.

    Args:
        slide: Target pptx Slide.
        table_shape_id: Table shape ID.
        mutations: List of mutation dicts with row, column, text, and optional style properties.

    Returns:
        Structured execution summary.
    """
    if not mutations:
        raise ValueError("Mutations list cannot be empty")

    shape = _find_table_shape(slide, table_shape_id)
    table = shape.table

    num_rows = len(table.rows)
    num_cols = len(table.columns)

    # Step 1: Pre-validation of all coordinates before mutation
    for idx, m in enumerate(mutations):
        if not isinstance(m, dict):
            raise ValueError(f"Mutation at index {idx} must be a dictionary")
        r = m.get("row")
        c = m.get("column")
        if r is None or c is None:
            raise ValueError(f"Mutation at index {idx} missing 'row' or 'column'")
        if r < 0 or r >= num_rows:
            raise IndexError(f"Row index {r} out of range (0..{num_rows - 1})")
        if c < 0 or c >= num_cols:
            raise IndexError(f"Column index {c} out of range (0..{num_cols - 1})")

    # Step 2: Apply mutations
    applied = []
    for m in mutations:
        r = m["row"]
        c = m["column"]
        cell = table.cell(r, c)

        # Update text
        if "text" in m and m["text"] is not None:
            new_text = str(m["text"])
            # If cell has paragraphs, preserve first run formatting if not overridden
            existing_font_name = None
            existing_font_size = None
            existing_bold = None
            existing_italic = None
            existing_color = None

            if cell.text_frame and cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                r0 = cell.text_frame.paragraphs[0].runs[0]
                existing_font_name = r0.font.name
                existing_font_size = r0.font.size
                existing_bold = r0.font.bold
                existing_italic = r0.font.italic
                if r0.font.color and getattr(r0.font.color, "rgb", None):
                    existing_color = r0.font.color.rgb

            cell.text = new_text

            # Re-apply formatting to paragraph / runs
            p = cell.text_frame.paragraphs[0]
            if p.runs:
                run = p.runs[0]
                # Font family
                f_name = m.get("font_name", existing_font_name)
                if f_name:
                    run.font.name = f_name
                # Font size
                if "font_size" in m and m["font_size"] is not None:
                    run.font.size = Pt(float(m["font_size"]))
                elif existing_font_size is not None:
                    run.font.size = existing_font_size
                # Bold
                if "bold" in m and m["bold"] is not None:
                    run.font.bold = bool(m["bold"])
                elif existing_bold is not None:
                    run.font.bold = existing_bold
                # Italic
                if "italic" in m and m["italic"] is not None:
                    run.font.italic = bool(m["italic"])
                elif existing_italic is not None:
                    run.font.italic = existing_italic
                # Font color
                if "font_color" in m and m["font_color"]:
                    run.font.color.rgb = _hex_to_rgb(m["font_color"])
                elif existing_color is not None:
                    run.font.color.rgb = existing_color

        # Horizontal alignment
        if "align" in m and m["align"]:
            align_lower = str(m["align"]).lower()
            align_map = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY,
            }
            if align_lower in align_map:
                for p in cell.text_frame.paragraphs:
                    p.alignment = align_map[align_lower]

        # Vertical alignment
        if "vertical_align" in m and m["vertical_align"]:
            valign_lower = str(m["vertical_align"]).lower()
            valign_map = {
                "top": MSO_ANCHOR.TOP,
                "middle": MSO_ANCHOR.MIDDLE,
                "center": MSO_ANCHOR.MIDDLE,
                "bottom": MSO_ANCHOR.BOTTOM,
            }
            if valign_lower in valign_map:
                cell.vertical_anchor = valign_map[valign_lower]

        # Cell fill
        if "fill" in m and m["fill"]:
            cell.fill.solid()
            cell.fill.fore_color.rgb = _hex_to_rgb(m["fill"])

        applied.append({"row": r, "column": c, "success": True})

    return {
        "success": True,
        "table_shape_id": table_shape_id,
        "mutations_applied": len(applied),
        "total_mutations": len(mutations),
        "applied": applied,
    }


def set_table_geometry(
    slide: Any,
    table_shape_id: int,
    left: Optional[float] = None,
    top: Optional[float] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    column_widths: Optional[List[float]] = None,
    row_heights: Optional[List[float]] = None,
) -> Dict[str, Any]:
    """Update table position, bounds, and individual column widths / row heights with PATCH semantics.

    Args:
        slide: Target pptx Slide.
        table_shape_id: Table shape ID.
        left: Left position in inches.
        top: Top position in inches.
        width: Total width in inches.
        height: Total height in inches.
        column_widths: List of column widths in inches.
        row_heights: List of row heights in inches.

    Returns:
        Dictionary containing updated geometry.
    """
    shape = _find_table_shape(slide, table_shape_id)
    table = shape.table

    if left is not None:
        shape.left = inches_to_emu(left)
    if top is not None:
        shape.top = inches_to_emu(top)
    if width is not None:
        if width <= 0:
            raise ValueError(f"Table width must be positive, got {width}")
        shape.width = inches_to_emu(width)
    if height is not None:
        if height <= 0:
            raise ValueError(f"Table height must be positive, got {height}")
        shape.height = inches_to_emu(height)

    if column_widths:
        if len(column_widths) > len(table.columns):
            raise ValueError(f"Got {len(column_widths)} column widths for table with {len(table.columns)} columns")
        for idx, w in enumerate(column_widths):
            if w <= 0:
                raise ValueError(f"Column width must be positive, got {w} at index {idx}")
            table.columns[idx].width = inches_to_emu(w)

    if row_heights:
        if len(row_heights) > len(table.rows):
            raise ValueError(f"Got {len(row_heights)} row heights for table with {len(table.rows)} rows")
        for idx, h in enumerate(row_heights):
            if h <= 0:
                raise ValueError(f"Row height must be positive, got {h} at index {idx}")
            table.rows[idx].height = inches_to_emu(h)

    return {
        "success": True,
        "table_shape_id": table_shape_id,
        "bbox": {
            "x": emu_to_inches(shape.left),
            "y": emu_to_inches(shape.top),
            "width": emu_to_inches(shape.width),
            "height": emu_to_inches(shape.height),
        },
        "column_widths": [emu_to_inches(c.width) for c in table.columns],
        "row_heights": [emu_to_inches(r.height) for r in table.rows],
    }


def _parse_cell_range(range_str: Optional[str], num_rows: int, num_cols: int) -> List[Tuple[int, int]]:
    """Parse range string into list of (row, col) tuples.

    Supported formats:
        - None or 'all': all cells
        - '0:0': single cell (row 0, col 0)
        - '0:0-1:2': rectangular block from (0,0) to (1,2) inclusive
        - 'row:0': entire row 0
        - 'col:1': entire column 1
    """
    if not range_str or range_str.lower() == "all":
        return [(r, c) for r in range(num_rows) for c in range(num_cols)]

    range_clean = range_str.strip().lower()

    # row:X
    if range_clean.startswith("row:"):
        r_idx = int(range_clean.split(":")[1])
        if r_idx < 0 or r_idx >= num_rows:
            raise IndexError(f"Row index {r_idx} out of range (0..{num_rows - 1})")
        return [(r_idx, c) for c in range(num_cols)]

    # col:X
    if range_clean.startswith("col:"):
        c_idx = int(range_clean.split(":")[1])
        if c_idx < 0 or c_idx >= num_cols:
            raise IndexError(f"Column index {c_idx} out of range (0..{num_cols - 1})")
        return [(r, c_idx) for r in range(num_rows)]

    # 0:0-1:2
    if "-" in range_clean:
        start_part, end_part = range_clean.split("-")
        s_r, s_c = map(int, start_part.split(":"))
        e_r, e_c = map(int, end_part.split(":"))
        if s_r < 0 or e_r >= num_rows or s_c < 0 or e_c >= num_cols or s_r > e_r or s_c > e_c:
            raise ValueError(f"Invalid range bounds: {range_str}")
        return [(r, c) for r in range(s_r, e_r + 1) for c in range(s_c, e_c + 1)]

    # single cell 0:0
    if ":" in range_clean:
        r, c = map(int, range_clean.split(":"))
        if r < 0 or r >= num_rows or c < 0 or c >= num_cols:
            raise IndexError(f"Cell ({r},{c}) out of range")
        return [(r, c)]

    raise ValueError(f"Unrecognized range format: {range_str}. Use '0:0', '0:0-1:2', 'row:0', 'col:1', or 'all'")


def _set_cell_border(cell: Any, side: str, color_hex: str, width_pt: float = 1.0) -> None:
    """Set OpenXML border on a table cell for a specific side."""
    side_tags = {"left": "lnL", "right": "lnR", "top": "lnT", "bottom": "lnB"}
    tag = side_tags.get(side.lower())
    if not tag:
        return
    clean_color = color_hex.strip().lstrip("#").upper()
    w_emu = pt_to_emu(width_pt)
    tcPr = cell._tc.get_or_add_tcPr()

    # Remove existing element if present
    for child in list(tcPr):
        if child.tag.endswith(tag):
            tcPr.remove(child)

    border_elem = parse_xml(
        f'<a:{tag} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{w_emu}" cmpd="s">'
        f'<a:solidFill><a:srgbClr val="{clean_color}"/></a:solidFill>'
        f'</a:{tag}>'
    )
    tcPr.append(border_elem)


def style_table(
    slide: Any,
    table_shape_id: int,
    range_spec: Optional[str] = None,
    style: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Apply styling (fill, fonts, alignment, margins, borders) to a table or cell range.

    Args:
        slide: Target pptx Slide.
        table_shape_id: Table shape ID.
        range_spec: Cell range e.g. 'all', '0:0', '0:0-1:2', 'row:0', 'col:1'.
        style: Dictionary containing style properties.

    Returns:
        Summary of styled cells.
    """
    if not style:
        raise ValueError("Style dictionary cannot be empty")

    shape = _find_table_shape(slide, table_shape_id)
    table = shape.table

    num_rows = len(table.rows)
    num_cols = len(table.columns)

    target_cells = _parse_cell_range(range_spec, num_rows, num_cols)

    for r, c in target_cells:
        cell = table.cell(r, c)

        # Fill
        if "fill" in style and style["fill"]:
            cell.fill.solid()
            cell.fill.fore_color.rgb = _hex_to_rgb(style["fill"])

        # Margins
        if "margins" in style and isinstance(style["margins"], dict):
            m = style["margins"]
            if "left" in m and m["left"] is not None:
                cell.margin_left = Inches(float(m["left"]))
            if "right" in m and m["right"] is not None:
                cell.margin_right = Inches(float(m["right"]))
            if "top" in m and m["top"] is not None:
                cell.margin_top = Inches(float(m["top"]))
            if "bottom" in m and m["bottom"] is not None:
                cell.margin_bottom = Inches(float(m["bottom"]))

        # Vertical alignment
        if "vertical_alignment" in style and style["vertical_alignment"]:
            valign_map = {
                "top": MSO_ANCHOR.TOP,
                "middle": MSO_ANCHOR.MIDDLE,
                "center": MSO_ANCHOR.MIDDLE,
                "bottom": MSO_ANCHOR.BOTTOM,
            }
            v_key = str(style["vertical_alignment"]).lower()
            if v_key in valign_map:
                cell.vertical_anchor = valign_map[v_key]

        # Paragraph & text styles
        if cell.text_frame:
            for p in cell.text_frame.paragraphs:
                # Horizontal alignment
                if "horizontal_alignment" in style and style["horizontal_alignment"]:
                    halign_map = {
                        "left": PP_ALIGN.LEFT,
                        "center": PP_ALIGN.CENTER,
                        "right": PP_ALIGN.RIGHT,
                        "justify": PP_ALIGN.JUSTIFY,
                    }
                    h_key = str(style["horizontal_alignment"]).lower()
                    if h_key in halign_map:
                        p.alignment = halign_map[h_key]

                # Runs styling
                for run in p.runs:
                    if "font_name" in style and style["font_name"]:
                        run.font.name = str(style["font_name"])
                    if "font_size" in style and style["font_size"] is not None:
                        run.font.size = Pt(float(style["font_size"]))
                    if "bold" in style and style["bold"] is not None:
                        run.font.bold = bool(style["bold"])
                    if "italic" in style and style["italic"] is not None:
                        run.font.italic = bool(style["italic"])
                    if "font_color" in style and style["font_color"]:
                        run.font.color.rgb = _hex_to_rgb(style["font_color"])

        # Borders
        if "borders" in style and isinstance(style["borders"], dict):
            b_info = style["borders"]
            b_color = b_info.get("color", "CCCCCC")
            b_width = float(b_info.get("width", 1.0))
            sides = b_info.get("sides", ["top", "bottom", "left", "right"])
            for s in sides:
                _set_cell_border(cell, s, b_color, b_width)

    return {
        "success": True,
        "table_shape_id": table_shape_id,
        "range": range_spec or "all",
        "styled_cells_count": len(target_cells),
    }


def merge_table_cells(
    slide: Any,
    table_shape_id: int,
    start_row: int,
    start_column: int,
    end_row: int,
    end_column: int,
) -> Dict[str, Any]:
    """Merge rectangular range of table cells.

    Args:
        slide: Target pptx Slide.
        table_shape_id: Table shape ID.
        start_row: Top row index (0-indexed).
        start_column: Left column index (0-indexed).
        end_row: Bottom row index (0-indexed).
        end_column: Right column index (0-indexed).

    Returns:
        Updated table inspection summary.
    """
    shape = _find_table_shape(slide, table_shape_id)
    table = shape.table

    num_rows = len(table.rows)
    num_cols = len(table.columns)

    if (
        start_row < 0
        or end_row >= num_rows
        or start_column < 0
        or end_column >= num_cols
        or start_row > end_row
        or start_column > end_column
    ):
        raise ValueError(
            f"Invalid merge range: ({start_row}, {start_column}) to ({end_row}, {end_column}). "
            f"Table size is {num_rows} rows x {num_cols} cols."
        )

    origin_cell = table.cell(start_row, start_column)
    corner_cell = table.cell(end_row, end_column)
    origin_cell.merge(corner_cell)

    return {
        "success": True,
        "table_shape_id": table_shape_id,
        "merged_range": {
            "start_row": start_row,
            "start_column": start_column,
            "end_row": end_row,
            "end_column": end_column,
        },
        "is_merge_origin": origin_cell.is_merge_origin,
    }