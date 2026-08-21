"""Data models for entire PowerPoint presentations and metadata."""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from powerpoint_mcp.models.slide import SlideModel


@dataclass
class PresentationMetadata:
    """Core document properties and metadata of a PowerPoint presentation."""
    title: Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
    created: Optional[str] = None
    modified: Optional[str] = None
    revision: Optional[int] = None
    category: Optional[str] = None
    comments: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Serialize metadata to dictionary."""
        return {
            "title": self.title,
            "author": self.author,
            "subject": self.subject,
            "created": self.created,
            "modified": self.modified,
            "revision": self.revision,
            "category": self.category,
            "comments": self.comments,
        }


@dataclass
class PresentationModel:
    """Model representing an entire PowerPoint presentation."""
    path: str
    width_inches: float
    height_inches: float
    width_emu: int
    height_emu: int
    slide_count: int
    theme_name: Optional[str] = None
    layouts: List[str] = field(default_factory=list)
    slides: List[SlideModel] = field(default_factory=list)
    slide_titles: List[Dict[str, Any]] = field(default_factory=list)
    metadata: PresentationMetadata = field(default_factory=PresentationMetadata)

    @property
    def presentation_path(self) -> str:
        """Alias for path."""
        return self.path

    @property
    def available_layouts(self) -> List[str]:
        """Alias for layouts."""
        return self.layouts

    def get_slide(self, slide_number: int) -> Optional[SlideModel]:
        """Retrieve a slide by its 1-indexed number."""
        if 1 <= slide_number <= len(self.slides):
            return self.slides[slide_number - 1]
        for slide in self.slides:
            if slide.slide_number == slide_number:
                return slide
        return None

    def to_dict(self) -> Dict[str, Any]:
        """Serialize presentation model to dictionary matching API specifications."""
        return {
            "path": self.path,
            "presentation_path": self.path,
            "slide_count": self.slide_count,
            "width_inches": self.width_inches,
            "height_inches": self.height_inches,
            "width_emu": self.width_emu,
            "height_emu": self.height_emu,
            "dimensions": {
                "width_inches": self.width_inches,
                "height_inches": self.height_inches,
                "width_emu": self.width_emu,
                "height_emu": self.height_emu,
            },
            "theme_name": self.theme_name,
            "layouts": self.layouts,
            "available_layouts": self.layouts,
            "slide_titles": self.slide_titles,
            "slides": [s.to_dict() for s in self.slides],
            "metadata": self.metadata.to_dict(),
        }
