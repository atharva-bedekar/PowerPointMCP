"""Rendering and visual verification subsystem for PowerPoint MCP Server."""

from powerpoint_mcp.rendering.image_diff import (
    VisualDiffResult,
    visual_diff,
)
from powerpoint_mcp.rendering.renderer import (
    BaseRenderer,
    LibreOfficeRenderer,
    NullRenderer,
    PowerPointRenderer,
    get_available_renderer,
)
from powerpoint_mcp.rendering.visual_compare import (
    SlideComparisonResult,
    compare_slides,
)

__all__ = [
    "BaseRenderer",
    "LibreOfficeRenderer",
    "NullRenderer",
    "PowerPointRenderer",
    "SlideComparisonResult",
    "VisualDiffResult",
    "compare_slides",
    "get_available_renderer",
    "visual_diff",
]
