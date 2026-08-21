# Project: PowerPoint MCP Server

## Architecture
Production-quality local Model Context Protocol (MCP) server in Python enabling deterministic PowerPoint (.pptx) inspection, precise editing, visual rendering, diffing, and rule-based validation for Antigravity.

```
powerpoint-mcp/
├── pyproject.toml
├── README.md
├── .gitignore
│
├── src/
│   └── powerpoint_mcp/
│       ├── __init__.py
│       ├── server.py
│       │
│       ├── models/
│       │   ├── __init__.py
│       │   ├── presentation.py
│       │   ├── slide.py
│       │   └── shape.py
│       │
│       ├── pptx/
│       │   ├── __init__.py
│       │   ├── inspector.py
│       │   ├── editor.py
│       │   ├── ooxml.py
│       │   ├── geometry.py
│       │   ├── styles.py
│       │   └── relationships.py
│       │
│       ├── rendering/
│       │   ├── __init__.py
│       │   ├── renderer.py
│       │   ├── image_diff.py
│       │   └── visual_compare.py
│       │
│       ├── tools/
│       │   ├── __init__.py
│       │   ├── inspection.py
│       │   ├── editing.py
│       │   ├── rendering.py
│       │   └── versioning.py
│       │
│       └── utils/
│           ├── __init__.py
│           ├── paths.py
│           ├── validation.py
│           └── logging.py
│
├── tests/
│   ├── __init__.py
│   ├── conftest.py
│   ├── fixtures/
│   │   ├── __init__.py
│   │   └── create_synthetic_deck.py
│   ├── test_inspection.py
│   ├── test_geometry.py
│   ├── test_editing.py
│   ├── test_text.py
│   ├── test_ooxml.py
│   ├── test_rendering.py
│   ├── test_validation.py
│   ├── test_session.py
│   ├── test_mcp.py
│   └── test_e2e_workflow.py
│
├── scripts/
│   ├── inspect_pptx.py
│   └── render_pptx.py
│
└── .agents/
    ├── skills/
    │   └── powerpoint-editor/
    │       └── SKILL.md
    └── mcp_config.json
```

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Presentation Metadata Inspection | `ppt_inspect_presentation`: dimensions, slide count, theme, layouts, titles, notes | M1 | Spec §3.1 |
| 2 | Slide & Shape Tree Inspection | `ppt_inspect_slide`: detailed shapes, EMU/inch coords, semantic roles, z-order | M1 | Spec §3.2 |
| 3 | Single Shape Deep Inspection | `ppt_inspect_shape`: deep properties, text frame, fill/line, OOXML details | M1 | Spec §3.3 |
| 4 | Semantic Shape Matching | `match_shapes(slide_a, slide_b)` multi-factor scoring & confidence | M1 | Spec §6 |
| 5 | Shape Geometry Modification | `ppt_modify_shape`: absolute/delta coordinates, dimensions, rotation, z-order | M2 | Spec §3.4 |
| 6 | Alignment & Distribution Geometry | `align_shapes_*`, `distribute_shapes_*`, `equalize_*`, collision math | M2 | Spec §7 |
| 7 | Text Editing & Run Style Preservation | `ppt_modify_text`: font, size, bold, italic, color, alignment, run-level styles | M2 | Spec §3.5 |
| 8 | Deterministic Copy/Move/Resize/Delete | `ppt_copy_shape`, `ppt_move_shape`, `ppt_resize_shape`, `ppt_delete_shape` | M2 | Spec §3.6 |
| 9 | OOXML Safe Fallback Helpers | `ppt_modify_ooxml`, `pptx/ooxml.py` for gradients, shadows, transparency | M2 | Spec §3.7 |
| 10 | PowerPoint COM Rendering Engine | Headless `PowerPoint.Application` COM export on Windows with clean lifecycle | M3 | Spec §3.8, §15 |
| 11 | LibreOffice Headless Rendering Engine | Headless `soffice` fallback rendering | M3 | Spec §3.8, §15 |
| 12 | Automatic Renderer Detection | `PPT_RENDERER=auto`, detection & reporting | M3 | Spec §15, §20 |
| 13 | Cross-Slide Geometric Comparison | `ppt_compare_slides`: geometry, text, dimensions, layout comparison | M3 | Spec §4 |
| 14 | Visual Image Diffing & Bounding Regions | `ppt_visual_diff`: pixel diff, changed bounding boxes, similarity score | M3 | Spec §4, §16 |
| 15 | Rule-Based Slide Validation | `ppt_validate_slide`: overlap, clipping, text overflow, tiny fonts, warnings | M4 | Spec §8 |
| 16 | Session & Working-Copy Architecture | `.ppt-agent/sessions/<id>/`, non-destructive working.pptx lifecycle | M4 | Spec §10 |
| 17 | Timestamped Backups & Revert | `ppt_create_backup`, `ppt_revert`, `ppt_save`, `ppt_save_as` | M4 | Spec §9, §10 |
| 18 | MCP Server Stdio & 19 Tools | `MCPServer` registration, typed schemas, structured errors | M5 | Spec §11, §25 |
| 19 | MCP Resources | `ppt://current/presentation`, `ppt://current/slide/{num}`, render resource | M5 | Spec §12 |
| 20 | Standalone CLI Tools | `scripts/inspect_pptx.py` and `scripts/render_pptx.py` | M5 | Spec §19 |
| 21 | Antigravity Packaging & Config | `.agents/mcp_config.json` and `.agents/skills/powerpoint-editor/SKILL.md` | M5 | Spec §13, §21 |
| 22 | Synthetic 3-Slide Presentation | Programmatic generator with feature boxes, 2-column, intentional overlaps | E2E | Spec §18 |
| 23 | Comprehensive Pytest Suite | Pytest suite covering all modules & in-memory MCP client | E2E | Spec §17 |
| 24 | Complete Documentation | `README.md` with architecture, setup, troubleshooting, conversational examples | M5 | Spec §22 |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| E2E | E2E Test Suite Track | Synthetic presentation, test runner, Tiers 1-4 test suite, `TEST_READY.md` | none | DONE |
| M1 | Core Models & PPTX Inspection | `models/*`, `pptx/inspector.py`, `pptx/styles.py`, semantic role inference, `match_shapes` | none | DONE |
| M2 | Geometry, Manipulation & OOXML | `pptx/geometry.py`, `pptx/editor.py`, `pptx/ooxml.py`, `pptx/relationships.py` | M1 | DONE |
| M3 | Rendering & Visual Verification | `rendering/renderer.py`, `rendering/image_diff.py`, `rendering/visual_compare.py` | M1 | DONE |
| M4 | Session, Safety & Validation | `utils/paths.py`, `utils/validation.py`, `tools/versioning.py`, working copies, backups | M1, M2 | DONE |
| M5 | MCP Server, CLI, Skill & Docs | `server.py`, `tools/*`, `scripts/*`, `mcp_config.json`, `SKILL.md`, `README.md` | M1, M2, M3, M4 | DONE |
| FINAL | Final Verification & Hardening | 100% E2E test pass, Tier 5 adversarial hardening, forensic audit | E2E, M5 | DONE |

## Interface Contracts

### `powerpoint_mcp.models`
- `BoundingBox`: `left_emu: int`, `top_emu: int`, `width_emu: int`, `height_emu: int`, inch properties (`left_inches`, `top_inches`, `width_inches`, `height_inches`, `right_inches`, `bottom_inches`), `to_dict()`.
- `TextStyle`: `font_name`, `font_size_pt`, `bold`, `italic`, `underline`, `color_rgb`, `alignment`, `line_spacing_pt`, `space_before_pt`, `space_after_pt`.
- `ShapeModel`: `shape_id: int`, `name: str`, `shape_type: ShapeType`, `role: SemanticRole`, `bbox: BoundingBox`, `rotation: float`, `z_order: int`, `text_frame: Optional[TextFrameModel]`, `fill: Dict[str, Any]`, `line: Dict[str, Any]`, `properties: Dict[str, Any]`, `to_dict()`.
- `SlideModel`: `slide_number: int`, `slide_id: int`, `title: Optional[str]`, `layout_name: str`, `width_inches: float`, `height_inches: float`, `shapes: List[ShapeModel]`, `notes: Optional[str]`, `to_dict()`.
- `PresentationModel`: `path: str`, `width_inches: float`, `height_inches: float`, `slide_count: int`, `theme_name: Optional[str]`, `layouts: List[str]`, `slides: List[SlideModel]`, `to_dict()`.

### `powerpoint_mcp.pptx.inspector`
- `inspect_presentation(path_or_prs: Union[str, Presentation]) -> PresentationModel`
- `inspect_slide(path_or_prs: Union[str, Presentation], slide_number: int) -> SlideModel`
- `inspect_shape(path_or_prs: Union[str, Presentation], slide_number: int, shape_id: int) -> ShapeModel`
- `infer_semantic_role(shape, slide_width_emu: int, slide_height_emu: int) -> SemanticRole`
- `match_shapes(slide_a: SlideModel, slide_b: SlideModel) -> List[Dict[str, Any]]`

### `powerpoint_mcp.pptx.geometry`
- `align_shapes(shapes: List[Shape], alignment: AlignmentType) -> None`
- `distribute_shapes(shapes: List[Shape], mode: DistributionMode, spacing: SpacingMode) -> None`
- `equalize_dimensions(shapes: List[Shape], equalize_width: bool, equalize_height: bool) -> None`
- `check_bounding_box_collision(b1: BoundingBox, b2: BoundingBox, tolerance_emu: int = 0) -> bool`
- `calculate_overlap_area(b1: BoundingBox, b2: BoundingBox) -> int`

### `powerpoint_mcp.pptx.editor`
- `modify_shape(slide, shape_id: int, x: Optional[float] = None, y: Optional[float] = None, width: Optional[float] = None, height: Optional[float] = None, rotation: Optional[float] = None, z_order: Optional[int] = None) -> Dict[str, Any]`
- `modify_text(slide, shape_id: int, text: Optional[str] = None, font_family: Optional[str] = None, font_size: Optional[float] = None, bold: Optional[bool] = None, italic: Optional[bool] = None, underline: Optional[bool] = None, color: Optional[str] = None, alignment: Optional[str] = None, paragraph_spacing: Optional[float] = None, line_spacing: Optional[float] = None, margins: Optional[Dict[str, float]] = None) -> Dict[str, Any]`
- `copy_shape(slide, shape_id: int, target_slide = None, offset_x_inches: float = 0.2, offset_y_inches: float = 0.2) -> int`
- `move_shape(slide, shape_id: int, delta_x_inches: Optional[float] = None, delta_y_inches: Optional[float] = None, x_inches: Optional[float] = None, y_inches: Optional[float] = None) -> Dict[str, Any]`
- `resize_shape(slide, shape_id: int, width_inches: Optional[float] = None, height_inches: Optional[float] = None, scale_x: Optional[float] = None, scale_y: Optional[float] = None) -> Dict[str, Any]`
- `delete_shape(slide, shape_id: int) -> bool`

### `powerpoint_mcp.rendering`
- `BaseRenderer.render_slide(presentation_path: str, slide_number: int, output_path: str, width: int = 1920, height: int = 1080) -> str`
- `BaseRenderer.render_presentation(presentation_path: str, output_dir: str) -> List[str]`
- `PowerPointRenderer`: COM automation with STA lifecycle management
- `LibreOfficeRenderer`: Headless soffice CLI fallback
- `get_available_renderer() -> BaseRenderer`
- `visual_diff(image_a_path: str, image_b_path: str, diff_output_path: Optional[str] = None) -> VisualDiffResult`
- `compare_slides(slide_a_model: SlideModel, slide_b_model: SlideModel, slide_a_img: Optional[str] = None, slide_b_img: Optional[str] = None) -> SlideComparisonResult`

### `powerpoint_mcp.utils.validation`
- `validate_slide(slide_model: SlideModel) -> SlideValidationResult` (rule codes VAL-01 to VAL-10: overlaps, boundaries, tiny fonts, text overflow, etc.)

### `powerpoint_mcp.tools.versioning` & Session Manager
- `SessionManager`: session creation in `.ppt-agent/sessions/<id>/`, working copy tracking, backup generation `presentation.backup-YYYYMMDD-HHMMSS.pptx`, `ppt_open`, `ppt_save`, `ppt_save_as`, `ppt_revert`.

## Code Layout
- Package: `src/powerpoint_mcp`
- Tests: `tests/`
- Scripts: `scripts/`
- Antigravity files: `.agents/mcp_config.json`, `.agents/skills/powerpoint-editor/SKILL.md`
