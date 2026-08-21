"""Data model for PowerPoint slides."""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Union

from powerpoint_mcp.models.shape import SemanticRole, ShapeModel, ShapeType


@dataclass
class SlideModel:
    """Model representing an individual slide and its contained shapes."""
    slide_number: int  # 1-indexed
    slide_id: int
    layout_name: str = "Custom"
    title: Optional[str] = None
    width_inches: float = 10.0
    height_inches: float = 5.625
    width_emu: int = 9144000
    height_emu: int = 5143500
    shapes: List[ShapeModel] = field(default_factory=list)
    notes: Optional[str] = None
    has_notes: bool = False
    notes_text: Optional[str] = None

    def __post_init__(self) -> None:
        """Synchronize notes and has_notes fields."""
        if self.notes_text is not None and self.notes is None:
            self.notes = self.notes_text
        elif self.notes is not None and self.notes_text is None:
            self.notes_text = self.notes
        if self.notes is not None and self.notes.strip():
            self.has_notes = True

    @property
    def shape_count(self) -> int:
        """Return the number of shapes on the slide."""
        return len(self.shapes)

    def get_shape_by_id(self, shape_id: int) -> Optional[ShapeModel]:
        """Find a shape on this slide by its unique ID."""
        for shape in self.shapes:
            if shape.shape_id == shape_id:
                return shape
        return None

    def get_shapes_by_role(self, role: Union[SemanticRole, str]) -> List[ShapeModel]:
        """Find all shapes on this slide with a given semantic role."""
        target_role = role.value if isinstance(role, SemanticRole) else str(role).lower()
        return [s for s in self.shapes if s.semantic_role.value == target_role]

    def get_shapes_by_type(self, shape_type: Union[ShapeType, str]) -> List[ShapeModel]:
        """Find all shapes on this slide of a given structural type."""
        target_type = shape_type.value if isinstance(shape_type, ShapeType) else str(shape_type).lower()
        return [s for s in self.shapes if s.shape_type.value == target_type]

    def to_dict(self) -> Dict[str, Any]:
        """Serialize slide to dictionary matching API specifications."""
        return {
            "slide_number": self.slide_number,
            "slide_id": self.slide_id,
            "layout_name": self.layout_name,
            "title": self.title,
            "shape_count": len(self.shapes),
            "shapes": [s.to_dict() for s in self.shapes],
            "has_notes": self.has_notes,
            "notes": self.notes,
            "notes_text": self.notes_text,
            "width_inches": self.width_inches,
            "height_inches": self.height_inches,
            "width_emu": self.width_emu,
            "height_emu": self.height_emu,
        }
