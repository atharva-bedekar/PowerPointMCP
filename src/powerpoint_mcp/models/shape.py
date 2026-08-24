"""Data models for shapes, bounding boxes, text styles, and text frames."""

from dataclasses import dataclass, field
from enum import Enum
import math
from typing import Any, Dict, List, Optional, Union

# Exact conversion factors per ECMA-376 Standard
EMU_PER_INCH = 914400
EMU_PER_POINT = 12700
EMU_PER_CM = 360000
POINTS_PER_INCH = 72


def inches_to_emu(inches: float) -> int:
    """Convert float inches to exact integer EMUs."""
    return int(round(inches * EMU_PER_INCH))


def emu_to_inches(emu: int, precision: int = 4) -> float:
    """Convert integer EMUs to float inches with rounding precision."""
    return round(float(emu) / EMU_PER_INCH, precision)


def pt_to_emu(pt: float) -> int:
    """Convert points to exact integer EMUs."""
    return int(round(pt * EMU_PER_POINT))


def emu_to_pt(emu: int, precision: int = 2) -> float:
    """Convert EMUs to points."""
    return round(float(emu) / EMU_PER_POINT, precision)


def apply_delta_inches(current_emu: int, delta_inches: float) -> int:
    """Apply a relative shift in inches to an existing EMU value without cumulative rounding error."""
    return current_emu + inches_to_emu(delta_inches)


class SemanticRole(str, Enum):
    """Inferred semantic role of a shape on a slide."""
    TITLE = "title"
    SLIDE_TITLE = "slide_title"
    SUBTITLE = "subtitle"
    SECTION_HEADER = "section_header"
    CARD = "card"
    CARD_TITLE = "card_title"
    BODY = "body"
    BULLET = "bullet"
    BADGE = "badge"
    METRIC = "metric"
    FOOTER = "footer"
    BACKGROUND = "background"
    CONNECTOR = "connector"
    ICON = "icon"
    IMAGE = "image"
    DIAGRAM = "diagram"
    TABLE = "table"
    CHART = "chart"
    UNKNOWN = "unknown"


class ShapeType(str, Enum):
    """Structural type classification of a PowerPoint shape."""
    AUTO_SHAPE = "auto_shape"
    TEXT_BOX = "text_box"
    PICTURE = "picture"
    GROUP = "group"
    TABLE = "table"
    CHART = "chart"
    CONNECTOR = "connector"
    MEDIA = "media"
    UNKNOWN = "unknown"


class AlignmentType(str, Enum):
    """Shape and paragraph alignment positions."""
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"
    JUSTIFY = "justify"


class DistributionMode(str, Enum):
    """Direction mode for distributing shapes."""
    HORIZONTAL = "horizontal"
    VERTICAL = "vertical"


class SpacingMode(str, Enum):
    """Spacing strategy for shape distribution."""
    EQUAL_GAPS = "equal_gaps"
    EQUAL_CENTERS = "equal_centers"


@dataclass
class BoundingBox:
    """Bounding box for a shape stored internally in integer EMUs."""
    left_emu: int
    top_emu: int
    width_emu: int
    height_emu: int

    @property
    def right_emu(self) -> int:
        """Right edge in integer EMUs."""
        return self.left_emu + self.width_emu

    @property
    def bottom_emu(self) -> int:
        """Bottom edge in integer EMUs."""
        return self.top_emu + self.height_emu

    @property
    def center_x_emu(self) -> int:
        """Horizontal center in integer EMUs."""
        return self.left_emu + self.width_emu // 2

    @property
    def center_y_emu(self) -> int:
        """Vertical center in integer EMUs."""
        return self.top_emu + self.height_emu // 2

    @property
    def left_inches(self) -> float:
        """Left coordinate in inches (rounded to 4 decimal places)."""
        return round(self.left_emu / EMU_PER_INCH, 4)

    @property
    def top_inches(self) -> float:
        """Top coordinate in inches (rounded to 4 decimal places)."""
        return round(self.top_emu / EMU_PER_INCH, 4)

    @property
    def width_inches(self) -> float:
        """Width in inches (rounded to 4 decimal places)."""
        return round(self.width_emu / EMU_PER_INCH, 4)

    @property
    def height_inches(self) -> float:
        """Height in inches (rounded to 4 decimal places)."""
        return round(self.height_emu / EMU_PER_INCH, 4)

    @property
    def right_inches(self) -> float:
        """Right coordinate in inches (rounded to 4 decimal places)."""
        return round(self.right_emu / EMU_PER_INCH, 4)

    @property
    def bottom_inches(self) -> float:
        """Bottom coordinate in inches (rounded to 4 decimal places)."""
        return round(self.bottom_emu / EMU_PER_INCH, 4)

    @property
    def center_x_inches(self) -> float:
        """Center X coordinate in inches (rounded to 4 decimal places)."""
        return round(self.center_x_emu / EMU_PER_INCH, 4)

    @property
    def center_y_inches(self) -> float:
        """Center Y coordinate in inches (rounded to 4 decimal places)."""
        return round(self.center_y_emu / EMU_PER_INCH, 4)

    @property
    def x(self) -> float:
        """Alias for left_inches."""
        return self.left_inches

    @property
    def y(self) -> float:
        """Alias for top_inches."""
        return self.top_inches

    @property
    def x_inches(self) -> float:
        """Alias for left_inches."""
        return self.left_inches

    @property
    def y_inches(self) -> float:
        """Alias for top_inches."""
        return self.top_inches

    @classmethod
    def from_inches(cls, left: float, top: float, width: float, height: float) -> "BoundingBox":
        """Construct BoundingBox from inch float values."""
        return cls(
            left_emu=inches_to_emu(left),
            top_emu=inches_to_emu(top),
            width_emu=inches_to_emu(width),
            height_emu=inches_to_emu(height),
        )

    @classmethod
    def from_emu(cls, left: int, top: int, width: int, height: int) -> "BoundingBox":
        """Construct BoundingBox from integer EMUs."""
        return cls(
            left_emu=int(left),
            top_emu=int(top),
            width_emu=int(width),
            height_emu=int(height),
        )

    def to_dict(self) -> Dict[str, Any]:
        """Serialize bounding box coordinates to dictionary."""
        return {
            "left_inches": self.left_inches,
            "top_inches": self.top_inches,
            "width_inches": self.width_inches,
            "height_inches": self.height_inches,
            "right_inches": self.right_inches,
            "bottom_inches": self.bottom_inches,
            "center_x_inches": self.center_x_inches,
            "center_y_inches": self.center_y_inches,
            "left_emu": self.left_emu,
            "top_emu": self.top_emu,
            "width_emu": self.width_emu,
            "height_emu": self.height_emu,
            "x": self.left_inches,
            "y": self.top_inches,
            "width": self.width_inches,
            "height": self.height_inches,
            "right": self.right_inches,
            "bottom": self.bottom_inches,
        }


@dataclass
class TextStyle:
    """Typographical styling properties for text runs and paragraphs."""
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color_rgb: Optional[str] = None  # Hex format: "RRGGBB" or "#RRGGBB"
    alignment: Optional[str] = None   # "left", "center", "right", "justify"
    line_spacing_pt: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        """Serialize text style to dictionary."""
        return {
            "font_name": self.font_name,
            "font_size_pt": self.font_size_pt,
            "bold": self.bold,
            "italic": self.italic,
            "underline": self.underline,
            "color_rgb": self.color_rgb,
            "alignment": self.alignment,
            "line_spacing_pt": self.line_spacing_pt,
            "space_before_pt": self.space_before_pt,
            "space_after_pt": self.space_after_pt,
        }


@dataclass
class TextRunModel:
    """Model representing a formatted text run inside a paragraph."""
    text: str
    style: TextStyle = field(default_factory=TextStyle)
    hyperlink_target: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Serialize text run to dictionary."""
        res: Dict[str, Any] = {
            "text": self.text,
            "style": self.style.to_dict(),
        }
        if self.hyperlink_target:
            res["hyperlink_target"] = self.hyperlink_target
        return res


@dataclass
class ParagraphModel:
    """Model representing a single paragraph with runs in a text frame."""
    text: str
    runs: List[TextRunModel] = field(default_factory=list)
    alignment: Optional[str] = None
    level: int = 0
    line_spacing_pt: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        """Serialize paragraph to dictionary."""
        return {
            "text": self.text,
            "runs": [r.to_dict() for r in self.runs],
            "alignment": self.alignment,
            "level": self.level,
            "line_spacing_pt": self.line_spacing_pt,
            "space_before_pt": self.space_before_pt,
            "space_after_pt": self.space_after_pt,
        }


@dataclass
class TextFrameModel:
    """Model representing the entire text frame of a shape."""
    text: str
    paragraphs: List[ParagraphModel] = field(default_factory=list)
    word_wrap: bool = True
    margin_left_inches: float = 0.1
    margin_right_inches: float = 0.1
    margin_top_inches: float = 0.05
    margin_bottom_inches: float = 0.05
    vertical_anchor: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Serialize text frame to dictionary."""
        return {
            "text": self.text,
            "paragraph_count": len(self.paragraphs),
            "paragraphs": [p.to_dict() for p in self.paragraphs],
            "word_wrap": self.word_wrap,
            "margins": {
                "left_inches": self.margin_left_inches,
                "right_inches": self.margin_right_inches,
                "top_inches": self.margin_top_inches,
                "bottom_inches": self.margin_bottom_inches,
            },
            "vertical_anchor": self.vertical_anchor,
        }


@dataclass
class ShapeModel:
    """Comprehensive model representing a shape on a PowerPoint slide."""
    shape_id: int
    name: str
    shape_type: ShapeType
    semantic_role: SemanticRole
    bbox: BoundingBox
    rotation: float = 0.0
    z_order: int = 0
    text_frame: Optional[TextFrameModel] = None
    fill: Dict[str, Any] = field(default_factory=dict)
    line: Dict[str, Any] = field(default_factory=dict)
    properties: Dict[str, Any] = field(default_factory=dict)
    group_id: Optional[int] = None
    image_metadata: Optional[Dict[str, Any]] = None
    table_metadata: Optional[Dict[str, Any]] = None
    chart_metadata: Optional[Dict[str, Any]] = None

    @property
    def id(self) -> int:
        """Alias for shape_id."""
        return self.shape_id

    @property
    def type(self) -> ShapeType:
        """Alias for shape_type."""
        return self.shape_type

    @property
    def role(self) -> SemanticRole:
        """Alias for semantic_role."""
        return self.semantic_role

    @property
    def bounds(self) -> BoundingBox:
        """Alias for bbox."""
        return self.bbox

    @property
    def fill_color(self) -> Optional[str]:
        """Extracted fill color hex if present."""
        return self.fill.get("color")

    @property
    def fill_type(self) -> Optional[str]:
        """Extracted fill type if present."""
        return self.fill.get("type")

    @property
    def line_color(self) -> Optional[str]:
        """Extracted line color hex if present."""
        return self.line.get("color")

    @property
    def line_width_pt(self) -> Optional[float]:
        """Extracted line width in points if present."""
        return self.line.get("width_pt")

    def to_dict(self) -> Dict[str, Any]:
        """Serialize shape to dictionary matching MCP and API specifications."""
        res: Dict[str, Any] = {
            "shape_id": self.shape_id,
            "id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type.value,
            "type": self.shape_type.value,
            "semantic_role": self.semantic_role.value,
            "role": self.semantic_role.value,
            "bbox": self.bbox.to_dict(),
            "x": self.bbox.left_inches,
            "y": self.bbox.top_inches,
            "width": self.bbox.width_inches,
            "height": self.bbox.height_inches,
            "right": self.bbox.right_inches,
            "bottom": self.bbox.bottom_inches,
            "rotation": round(self.rotation, 2),
            "z_order": self.z_order,
            "fill": self.fill,
            "fill_color": self.fill_color,
            "line": self.line,
            "line_color": self.line_color,
            "properties": self.properties,
        }
        if self.group_id is not None:
            res["group_id"] = self.group_id

        if self.text_frame:
            res["text_frame"] = self.text_frame.to_dict()
            res["text"] = self.text_frame.text
            # Extract dominant text style from first non-empty paragraph/run
            first_style: Optional[TextStyle] = None
            alignment: Optional[str] = None
            for p in self.text_frame.paragraphs:
                if p.runs:
                    first_style = p.runs[0].style
                    alignment = p.alignment
                    break
            if first_style:
                res["font_family"] = first_style.font_name
                res["font_name"] = first_style.font_name
                res["font_size"] = first_style.font_size_pt
                res["font_size_pt"] = first_style.font_size_pt
                res["bold"] = first_style.bold
                res["italic"] = first_style.italic
                res["underline"] = first_style.underline
                res["color"] = first_style.color_rgb
                res["color_rgb"] = first_style.color_rgb
                res["alignment"] = alignment

        if self.image_metadata:
            res["image_metadata"] = self.image_metadata
        if self.table_metadata:
            res["table_metadata"] = self.table_metadata
        if self.chart_metadata:
            res["chart_metadata"] = self.chart_metadata

        return res

    def to_summary_dict(self) -> Dict[str, Any]:
        """Serialize shape to concise agent-friendly summary dictionary."""
        res: Dict[str, Any] = {
            "shape_id": self.shape_id,
            "id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type.value,
            "type": self.shape_type.value,
            "semantic_role": self.semantic_role.value,
            "role": self.semantic_role.value,
            "bbox": {
                "left_inches": self.bbox.left_inches,
                "top_inches": self.bbox.top_inches,
                "width_inches": self.bbox.width_inches,
                "height_inches": self.bbox.height_inches,
            },
            "x": self.bbox.left_inches,
            "y": self.bbox.top_inches,
            "width": self.bbox.width_inches,
            "height": self.bbox.height_inches,
            "right": self.bbox.right_inches,
            "bottom": self.bbox.bottom_inches,
            "rotation": round(self.rotation, 2),
            "z_order": self.z_order,
        }


        if self.fill_color:
            res["fill_color"] = self.fill_color
        if self.line_color:
            res["line_color"] = self.line_color
        if self.line_width_pt is not None:
            res["line_width_pt"] = self.line_width_pt

        if self.text_frame and self.text_frame.text:
            raw_text = self.text_frame.text.strip()
            summary_text = (raw_text[:120] + "...") if len(raw_text) > 120 else raw_text
            res["text"] = summary_text
            res["text_summary"] = summary_text

            for p in self.text_frame.paragraphs:
                if p.runs:
                    first_style = p.runs[0].style
                    if first_style.font_name:
                        res["font_family"] = first_style.font_name
                        res["font_name"] = first_style.font_name
                    if first_style.font_size_pt is not None:
                        res["font_size"] = first_style.font_size_pt
                        res["font_size_pt"] = first_style.font_size_pt
                    if first_style.bold is not None:
                        res["bold"] = first_style.bold
                    if first_style.color_rgb:
                        res["color"] = first_style.color_rgb
                        res["color_rgb"] = first_style.color_rgb
                    if p.alignment:
                        res["alignment"] = p.alignment
                    break

        if self.table_metadata:
            res["table"] = {
                "rows": self.table_metadata.get("rows", 0),
                "columns": self.table_metadata.get("columns", 0),
            }
        if self.chart_metadata:
            res["chart"] = {
                "chart_type": self.chart_metadata.get("chart_type", "unknown"),
            }

        return res

