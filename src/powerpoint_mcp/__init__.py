"""PowerPoint MCP Server package.

Provides deterministic inspection, editing, rendering, and validation for PowerPoint (.pptx) presentations.
"""

from powerpoint_mcp.models import (
    AlignmentType,
    BoundingBox,
    DistributionMode,
    EMU_PER_CM,
    EMU_PER_INCH,
    EMU_PER_POINT,
    ParagraphModel,
    POINTS_PER_INCH,
    PresentationMetadata,
    PresentationModel,
    SemanticRole,
    ShapeModel,
    ShapeType,
    SlideModel,
    SpacingMode,
    TextFrameModel,
    TextRunModel,
    TextStyle,
    apply_delta_inches,
    emu_to_inches,
    emu_to_pt,
    inches_to_emu,
    pt_to_emu,
)
from powerpoint_mcp.pptx import (
    PPTXInspector,
    infer_semantic_role,
    inspect_presentation,
    inspect_shape,
    inspect_slide,
    match_shapes,
)

__version__ = "0.1.0"

__all__ = [
    "AlignmentType",
    "BoundingBox",
    "DistributionMode",
    "EMU_PER_CM",
    "EMU_PER_INCH",
    "EMU_PER_POINT",
    "PPTXInspector",
    "ParagraphModel",
    "POINTS_PER_INCH",
    "PresentationMetadata",
    "PresentationModel",
    "SemanticRole",
    "ShapeModel",
    "ShapeType",
    "SlideModel",
    "SpacingMode",
    "TextFrameModel",
    "TextRunModel",
    "TextStyle",
    "__version__",
    "apply_delta_inches",
    "emu_to_inches",
    "emu_to_pt",
    "inches_to_emu",
    "infer_semantic_role",
    "inspect_presentation",
    "inspect_shape",
    "inspect_slide",
    "match_shapes",
    "pt_to_emu",
]
