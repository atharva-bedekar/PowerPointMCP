"""Data models for semantic components in PowerPoint slides."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Union

from powerpoint_mcp.models.shape import BoundingBox, EMU_PER_INCH, inches_to_emu


class ComponentType(str, Enum):
    """Semantic classification of a multi-shape PowerPoint component."""
    HEADER = "header"
    FOOTER = "footer"
    STEPPER = "stepper"
    CARD = "card"
    CARD_LIST = "card_list"
    CONTENT_CONTAINER = "content_container"
    CONTENT_AREA = "content_area"
    TITLE_BLOCK = "title_block"
    BADGE = "badge"
    METRIC_GROUP = "metric_group"
    CUSTOM = "custom"


@dataclass
class ComponentModel:
    """Logical model representing a visual component composed of one or more PowerPoint shapes."""
    component_id: str
    component_type: Union[ComponentType, str]
    slide_number: int
    shape_ids: List[int] = field(default_factory=list)
    bbox: Optional[BoundingBox] = None
    properties: Dict[str, Any] = field(default_factory=dict)
    children: List[Dict[str, Any]] = field(default_factory=list)
    confidence: float = 1.0

    @property
    def type_str(self) -> str:
        """String representation of component type."""
        if isinstance(self.component_type, ComponentType):
            return self.component_type.value
        return str(self.component_type).lower()

    def to_dict(self, detail: str = "summary") -> Dict[str, Any]:
        """Serialize component model to dictionary.

        Args:
            detail: 'summary' (concise high-level data) or 'full' (including children and raw properties).
        """
        bbox_dict = self.bbox.to_dict() if self.bbox else {
            "left": 0.0,
            "top": 0.0,
            "width": 0.0,
            "height": 0.0,
        }

        res: Dict[str, Any] = {
            "id": self.component_id,
            "type": self.type_str,
            "slide_number": self.slide_number,
            "shape_ids": self.shape_ids,
            "bbox": bbox_dict,
            "confidence": round(self.confidence, 2),
        }

        # Include key properties based on component type
        if self.type_str == "stepper":
            res["steps"] = self.properties.get("steps", [])
            res["active_step"] = self.properties.get("active_step")
            res["step_count"] = len(self.properties.get("steps", []))
            if "direction" in self.properties:
                res["direction"] = self.properties["direction"]
        elif self.type_str == "header" or self.type_str == "title_block":
            res["title"] = self.properties.get("title", "")
            if "subtitle" in self.properties:
                res["subtitle"] = self.properties.get("subtitle", "")
        elif self.type_str == "footer":
            res["footer_text"] = self.properties.get("text", "")
        elif self.type_str in ("card", "card_list", "content_container", "content_area"):
            if "title" in self.properties:
                res["title"] = self.properties["title"]
            if "item_count" in self.properties:
                res["item_count"] = self.properties["item_count"]
            if "items" in self.properties and detail == "full":
                res["items"] = self.properties["items"]

        if detail == "full":
            res["properties"] = self.properties
            res["children"] = self.children

        return res
