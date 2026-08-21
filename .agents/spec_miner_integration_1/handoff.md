# Technical Specification Handoff: R3 (Session & Safety Layer) & R4 (Antigravity Integration & Tooling)

**Author:** Integration & Safety Spec Miner  
**Date:** 2026-08-21T05:58:00Z  
**Target:** PowerPoint MCP Server  
**Working Directory:** `C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_integration_1`  
**Integrity Mode:** Development  

---

# 1. Observation

Based on direct inspection of authoritative specification documents and local system environment:
1. **ORIGINAL_REQUEST.md**:
   - Lines 18-20: Requirement R3 (Session & Safety Layer) mandates a non-destructive session/working-copy model with automatic timestamped backups before modifications (`ppt_save`, `ppt_save_as`, `ppt_revert`) to ensure source files are never overwritten unintentionally. Includes rule-based slide validation (`ppt_validate_slide`) for overlaps, boundary clipping, and text overflow.
   - Lines 21-23: Requirement R4 (Antigravity Integration & Tooling) mandates standard MCP stdio server tools (19 tools) and resources (`ppt://current/...`), standalone CLI debugging scripts (`inspect_pptx.py`, `render_pptx.py`), Antigravity workspace configuration (`.agents/mcp_config.json`), and an Antigravity skill (`.agents/skills/powerpoint-editor/SKILL.md`) guiding agent behavior through inspect-modify-render-verify loops.
   - Lines 26-44: Acceptance criteria demanding passing automated test suites with in-memory MCP client testing, synthetic 3-slide test presentation generation, end-to-end multi-step workflow verification, safe stdio MCP configuration, and a complete `README.md`.
2. **Build a PowerPoint Editing MCP Server for Antigravity.md**:
   - Section 1 (Architecture, lines 23-95): Defines project layout, Python 3.10+, FastMCP / MCP Python SDK, `uv` dependency management.
   - Section 3 & 11 (MCP Tool Design, lines 147-410, 643-680): Enumerates the 19 core tools: `ppt_open`, `ppt_inspect_presentation`, `ppt_inspect_slide`, `ppt_inspect_shape`, `ppt_modify_shape`, `ppt_modify_text`, `ppt_copy_shape`, `ppt_move_shape`, `ppt_resize_shape`, `ppt_delete_shape`, `ppt_modify_ooxml`, `ppt_validate_slide`, `ppt_render_slide`, `ppt_render_presentation`, `ppt_compare_slides`, `ppt_visual_diff`, `ppt_save`, `ppt_save_as`, `ppt_revert`.
   - Section 8 (Slide Validation, lines 549-580): Specifies rule-based slide validation detecting overlaps, off-slide objects, text overflow, tiny fonts, inconsistent title positions/margins, duplicate objects, extreme rotations, and reporting structured diagnostic warnings.
   - Section 9 & 10 (Versioning & Working-Copy Model, lines 581-642): Mandates `.ppt-agent/` directory structure (`sessions/<session-id>/working.pptx`, `backups/`, `renders/`, `diffs/`, `metadata.json`), timestamped backups (`presentation.backup-YYYYMMDD-HHMMSS.pptx`), and non-destructive save operations.
   - Section 12 (Resources, lines 681-696): Exposes `ppt://current/presentation`, `ppt://current/slide/{slide_number}`, `ppt://current/slide/{slide_number}/render`.
   - Section 13 & 14 (Antigravity PowerPoint Skill, lines 697-796): Requires `.agents/skills/powerpoint-editor/SKILL.md` with YAML frontmatter, 15 editing rules, decision trees, and inspect-reason-modify-render-verify workflow.
   - Section 17 & 18 (Tests & Synthetic Presentation, lines 849-925): Programmatic 3-slide test deck (Slide 1: title/subtitle/3 boxes/image placeholder, Slide 2: title/2-column/diagram/footer, Slide 3: title/intentional overlaps). In-memory FastMCP client tests for all 19 tools.
   - Section 19 (CLI Utilities, lines 927-947): `scripts/inspect_pptx.py` and `scripts/render_pptx.py`.
   - Section 20 & 21 (Configuration & MCP Config, lines 949-993): Environment variables (`PPT_RENDERER`, `PPT_WORKSPACE_DIR`, `PPT_BACKUP_ENABLED`, `PPT_DEFAULT_OUTPUT_DIR`), `.agents/mcp_config.json` with stdio transport, safe config merging.
   - Section 22 (README, lines 994-1028): 12 mandatory sections and conversational examples.
   - Section 25 (Error Handling, lines 1068-1090): Structured error handling returning informative diagnostic JSON instead of raw exceptions.
3. **Local System Probe**:
   - OS: Windows 11 Enterprise (x86_64).
   - Package manager: `uv` 0.10.11.
   - Python runtimes: Python 3.12.10, 3.13.9, 3.14.3 installed locally.
   - Microsoft PowerPoint: Installed at `C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE` (COM automation available via `win32com.client`).
   - LibreOffice: Not installed on host (fallback detection must gracefully report absence).

---

# 2. Logic Chain

1. **Session & Safety Architecture**:
   - To prevent destructive overwrites, when a user or agent invokes `ppt_open`, a unique UUID session is provisioned under `.ppt-agent/sessions/<session_id>/`.
   - The source presentation is copied into the session as `working.pptx`.
   - Before any mutating operation (`ppt_modify_shape`, `ppt_modify_text`, `ppt_delete_shape`, `ppt_modify_ooxml`, `ppt_save`), a timestamped snapshot is saved to `backups/presentation.backup-YYYYMMDD-HHMMSS.pptx`.
   - `ppt_revert` can instantly roll back to any prior snapshot or the original source file.
   - `ppt_save_as` copies the working copy to the user-specified destination without modifying the original source.
2. **Rule-Based Slide Validation (`ppt_validate_slide`)**:
   - Rather than relying on non-deterministic LLM visual appraisal, mathematical geometric algorithms calculate bounding box intersections in native EMUs / inches.
   - Overlaps are detected using Axis-Aligned Bounding Box (AABB) intersection tests, with an exclusion filter for background shape containers and a configurable minimum tolerance (>0.01 inches) to prevent false positives from adjacent snapped borders.
   - Boundary clipping detects shapes with negative coordinates or shapes extending past `slide_width` and `slide_height`.
   - Text overflow uses a heuristic estimator calculating text area (glyph count * avg char width * line height) against the text frame's available bounding box (accounting for inner margins).
   - Font sizes < 8pt or non-standard rotations (angles other than 0, 90, 180, 270 degrees) trigger structured warnings.
3. **MCP Server & Tool Interface (19 Tools)**:
   - Built on `mcp.server.fastmcp.FastMCP` over standard input/output (stdio).
   - All tools return typed Pydantic models serialized to clean JSON dictionaries.
   - Coordinates are consistently presented in decimal inches with high precision (e.g., 4 decimal places) while internally manipulated in EMUs (1 inch = 914,400 EMUs) to eliminate rounding drift.
   - Structured error handling intercepts custom exceptions (`ShapeNotFound`, `SlideNotFound`, `SessionNotFound`, `RendererError`, `OOXMLValidationError`) and converts them into informative structured JSON payloads containing available recovery options (e.g. list of valid shape IDs).
4. **Antigravity Skill & Configuration Integration**:
   - `.agents/mcp_config.json` configures the stdio server invocation using `uv run python -m powerpoint_mcp.server`.
   - If `.agents/mcp_config.json` already exists in a user workspace, a merge utility safely injects the `powerpoint-mcp` entry under `"mcpServers"` without wiping existing configurations.
   - `.agents/skills/powerpoint-editor/SKILL.md` implements Antigravity YAML frontmatter (`name: powerpoint-editor`, `description: ...`) and encodes the 15 immutable editing rules, inspect-reason-modify-render-verify decision trees, and batching heuristics to minimize token consumption and MCP call volume.
5. **Testing & Validation Strategy**:
   - Programmatic generation of a synthetic 3-slide `.pptx` (`tests/fixtures/sample_presentation.pptx`) covers clean multi-box layouts, 2-column bulleted architectures, and intentional overlaps/boundary violations.
   - In-memory MCP client tests using the FastMCP test harness verify all 19 tools and resources without process-spawning latency.

---

# 3. Caveats

1. **COM Automation in Headless/Agent Environments**:
   - On Windows, PowerPoint COM automation (`win32com.client`) requires PowerPoint to run in invisible mode (`Visible = False` / `WithWindow = False`). COM objects and presentation handles must be released inside robust `try...finally` blocks to prevent orphaned `POWERPNT.EXE` background processes.
2. **LibreOffice Absence**:
   - LibreOffice is not installed on the current host machine. The renderer detection mechanism in `powerpoint_mcp/rendering/renderer.py` must prioritize PowerPoint COM and return a clean, descriptive error/status message if LibreOffice fallback is requested when LibreOffice binary is absent.
3. **Text Overflow Heuristics**:
   - Python-pptx and headless environments without an active typography rendering engine cannot measure exact kerning and line breaks for arbitrary proprietary fonts. The text overflow validation heuristic provides conservative estimates based on character count, average glyph aspect ratios, line counts, font point sizes, and inner margin dimensions.

---

# 4. Conclusion

The specification for R3 (Session & Safety Layer) and R4 (Antigravity Integration & Tooling) is fully defined and ready for implementation. The architecture establishes a deterministic, non-destructive, highly structured editing pipeline that prevents data loss, isolates working state, validates geometric integrity, exposes 19 structured FastMCP tools and 3 resources, provides standalone CLI utilities, and provides seamless Antigravity integration via `.agents/mcp_config.json` and `.agents/skills/powerpoint-editor/SKILL.md`.

---

# 5. Verification Method

To verify the implementation against this specification:
1. **In-Memory MCP Client Tests**:
   - Run `uv run pytest tests/test_mcp.py` to verify tool discovery, schema validity, and tool execution for all 19 tools.
2. **Session & Safety Verification**:
   - Run `uv run pytest tests/test_session.py` to verify working-copy creation in `.ppt-agent/sessions/`, backup creation with timestamp patterns, `ppt_save`, `ppt_save_as`, and `ppt_revert`.
3. **Slide Validation Verification**:
   - Run `uv run pytest tests/test_validation.py` against the synthetic test presentation Slide 3 to verify detection of intentional overlaps, off-slide boundaries, and tiny fonts.
4. **CLI Utility Verification**:
   - Run `uv run python scripts/inspect_pptx.py tests/fixtures/sample_presentation.pptx` (verify structured output and exit code 0).
   - Run `uv run python scripts/render_pptx.py tests/fixtures/sample_presentation.pptx --output .ppt-agent/renders/` (verify PNG creation and exit code 0).
5. **Antigravity Customization Verification**:
   - Verify `.agents/mcp_config.json` is valid JSON and parses under the Antigravity MCP schema.
   - Verify `.agents/skills/powerpoint-editor/SKILL.md` contains valid YAML frontmatter and adheres to the Antigravity progressive disclosure standard.


## Features Discovered

| # | Category | Feature | Description | Inputs | Outputs | Error Behavior | Discovered Via |
|---|----------|---------|-------------|--------|---------|----------------|----------------|
| 1 | R3: Session | Session Working Copy | Provisions isolated directory in `.ppt-agent/sessions/<id>/` with `working.pptx` and `metadata.json` | `presentation_path: str` | `{"session_id": str, "working_path": str, "slide_count": int, "metadata": dict}` | `FileNotFoundError`, `FileAccessError` | Spec Section 9, 10 |
| 2 | R3: Safety | Timestamped Backups | Creates timestamped snapshot (`presentation.backup-YYYYMMDD-HHMMSS.pptx`) before mutating edits | `session_id: str` or `presentation_path: str` | `{"backup_path": str, "timestamp": str}` | `BackupError` if disk write fails | Spec Section 9 |
| 3 | R3: Session | `ppt_open` | Initializes or connects to a presentation editing session | `presentation_path: str` | Session info, metadata, slide summary | `FileNotFoundError` | Spec Section 10, 11 |
| 4 | R3: Session | `ppt_save` | Commits session changes to original file with prior backup | `session_id: Optional[str]`, `presentation_path: Optional[str]` | `{"saved_path": str, "backup_path": str, "timestamp": str}` | `SessionNotFound`, `FileAccessError` | Spec Section 9, 10 |
| 5 | R3: Session | `ppt_save_as` | Saves session working copy to new destination path | `output_path: str`, `overwrite: Optional[bool]` | `{"saved_path": str, "success": bool}` | `FileExistsError` (if overwrite=False) | Spec Section 9, 10 |
| 6 | R3: Session | `ppt_revert` | Reverts working copy to original or specified backup timestamp | `target: Optional[str]` ("original" or timestamp) | `{"reverted_to": str, "success": bool}` | `BackupNotFound`, `SessionNotFound` | Spec Section 10, 11 |
| 7 | R3: Validation | Overlap Detection | AABB geometric collision check between non-container shapes | `slide_number: int`, `threshold_inches: float` | List of overlapping shape pairs with overlap depth and area | `SlideNotFound` | Spec Section 7, 8 |
| 8 | R3: Validation | Off-Slide / Clipping | Detects shapes protruding beyond slide canvas (left, top, right, bottom) | `slide_number: int` | Protrusion report per shape in inches | `SlideNotFound` | Spec Section 8 |
| 9 | R3: Validation | Text Overflow Heuristic | Heuristic check of text volume vs text frame inner dimensions | `slide_number: int` | List of shapes with estimated overflow ratio | `SlideNotFound` | Spec Section 8 |
| 10 | R3: Validation | Suspicious Font Size | Flags font sizes below minimum threshold (< 8pt / < 10pt) | `slide_number: int`, `min_pt: float` | List of shapes/runs with tiny text | `SlideNotFound` | Spec Section 8 |
| 11 | R3: Validation | Inconsistent Title Check | Verifies title position and margins consistency across slides | `slide_number: int` | Title geometry variance vs slide master or slide 1 | `SlideNotFound` | Spec Section 8 |
| 12 | R3: Validation | Duplicate Object Detection | Identifies superimposed shapes with identical geometry and properties | `slide_number: int` | List of duplicate shape ID clusters | `SlideNotFound` | Spec Section 8 |
| 13 | R3: Validation | Extreme Rotation Check | Flags shapes rotated to non-cardinal or irregular angles | `slide_number: int` | List of shapes with irregular rotation angles | `SlideNotFound` | Spec Section 8 |
| 14 | R3: Validation | `ppt_validate_slide` | Combined validation tool returning structured warnings and summary | `slide_number: int`, `rules: Optional[List[str]]` | `{"valid": bool, "warning_count": int, "warnings": [...]}` | `SlideNotFound` | Spec Section 8, 11 |
| 15 | R4: MCP Tool | `ppt_inspect_presentation` | Returns presentation dimensions, slide count, theme, layout names, titles | `presentation_path: Optional[str]` | Structured presentation metadata dictionary | `PresentationNotFound` | Spec Section 3.1 |
| 16 | R4: MCP Tool | `ppt_inspect_slide` | Returns all shapes on slide with roles, geometry, typography, and fills | `slide_number: int` | Structured list of shape dictionaries with semantic roles | `SlideNotFound` | Spec Section 3.2 |
| 17 | R4: MCP Tool | `ppt_inspect_shape` | Returns complete details and OOXML snippet for a single shape | `slide_number: int`, `shape_id: int` | Detailed shape model dictionary | `ShapeNotFound`, `SlideNotFound` | Spec Section 3.3 |
| 18 | R4: MCP Tool | `ppt_modify_shape` | Modifies shape coordinates, dimensions, rotation, z-order, alignment | `slide_number: int`, `shape_id: int`, geometry kwargs | Updated shape geometry dictionary | `ShapeNotFound`, `InvalidCoordinate` | Spec Section 3.4 |
| 19 | R4: MCP Tool | `ppt_modify_text` | Modifies text content, typography, colors, and margins with run preservation | `slide_number: int`, `shape_id: int`, text kwargs | Updated text frame details dictionary | `ShapeNotFound`, `InvalidTextFrame` | Spec Section 3.5 |
| 20 | R4: MCP Tool | `ppt_copy_shape` | Duplicates shape to current or target slide with formatting preserved | `slide_number: int`, `shape_id: int`, offset kwargs | `{"new_shape_id": int, "shape": dict}` | `ShapeNotFound`, `SlideNotFound` | Spec Section 3.6 |
| 21 | R4: MCP Tool | `ppt_move_shape` | Moves shape by absolute (x, y) or relative delta (dx, dy) | `slide_number: int`, `shape_id: int`, dx, dy, x, y | Updated position dictionary | `ShapeNotFound` | Spec Section 3.6 |
| 22 | R4: MCP Tool | `ppt_resize_shape` | Resizes shape with absolute dimensions or scaling factors | `slide_number: int`, `shape_id: int`, width, height, scale | Updated dimension dictionary | `ShapeNotFound` | Spec Section 3.6 |
| 23 | R4: MCP Tool | `ppt_delete_shape` | Removes shape cleanly from slide shape collection | `slide_number: int`, `shape_id: int` | `{"deleted_shape_id": int, "remaining_shapes": [...]}` | `ShapeNotFound` | Spec Section 3.6 |
| 24 | R4: MCP Tool | `ppt_modify_ooxml` | Controlled XML modification helper for advanced properties | `slide_number: int`, `shape_id: Optional[int]`, op, xpath, xml | `{"success": bool, "xml_snippet": str}` | `OOXMLValidationError` | Spec Section 3.7 |
| 25 | R4: MCP Tool | `ppt_render_slide` | Renders slide to PNG via PowerPoint COM or LibreOffice | `slide_number: int`, `output_dir: Optional[str]`, `dpi: int` | `{"image_path": str, "renderer": str, "width": int, "height": int}` | `RendererError`, `SlideNotFound` | Spec Section 3.8, 15 |
| 26 | R4: MCP Tool | `ppt_render_presentation` | Renders all slides in presentation to PNGs in directory | `output_dir: Optional[str]`, `dpi: int` | `{"image_paths": List[str], "renderer": str}` | `RendererError` | Spec Section 3.8, 15 |
| 27 | R4: MCP Tool | `ppt_compare_slides` | Compares geometric and visual properties between two slides | `slide_a: int`, `slide_b: int`, `match_shapes: bool` | Geometric diffs, style differences, shape match confidences | `SlideNotFound` | Spec Section 4, 6 |
| 28 | R4: MCP Tool | `ppt_visual_diff` | Computes pixel diff, bounding boxes of change, and similarity score | `before_image: str`, `after_image: str`, `threshold: float` | `{"diff_image_path": str, "similarity_score": float, "changed_boxes": [...]}` | `FileNotFoundError` | Spec Section 4, 16 |
| 29 | R4: Resources | `ppt://current/presentation` | Exposes presentation metadata summary resource | None | JSON presentation summary | `SessionNotFound` | Spec Section 12 |
| 30 | R4: Resources | `ppt://current/slide/{N}` | Exposes structured slide content resource | `slide_number: int` | JSON slide description | `SlideNotFound` | Spec Section 12 |
| 31 | R4: Resources | `ppt://current/slide/{N}/render` | Exposes rendered slide image resource | `slide_number: int` | Image/PNG content or URI | `RendererError` | Spec Section 12 |
| 32 | R4: CLI Utility | `scripts/inspect_pptx.py` | Standalone CLI to inspect presentation/slide/shape in text or JSON | CLI flags: `presentation`, `--slide`, `--shape`, `--json` | Formatted stdout, exit code 0/1 | CLI Argument Error | Spec Section 19 |
| 33 | R4: CLI Utility | `scripts/render_pptx.py` | Standalone CLI to render presentation slides to PNG | CLI flags: `presentation`, `--slide`, `--output`, `--renderer`, `--dpi` | Output paths on stdout, exit code 0/1 | Renderer Failure | Spec Section 19 |
| 34 | R4: Customization | `.agents/mcp_config.json` | Configures stdio FastMCP server launch command for Antigravity | Workspace configuration file | JSON configuration | JSON Parse Error | Spec Section 21 |
| 35 | R4: Skill | `.agents/skills/powerpoint-editor/SKILL.md` | Antigravity skill with 15 rules, decision trees, inspect-modify-render loop | Antigravity skill markdown file | Skill instructions and triggers | None | Spec Section 13, 14 |

## Edge Cases

| # | Feature | Input | Observed Behavior |
|---|---------|-------|-------------------|
| 1 | `ppt_open` | Non-existent `.pptx` file path | Returns structured `FileNotFoundError` error JSON with `error_type: "FileNotFound"`, path searched, and success=False. |
| 2 | `ppt_open` | Corrupted / non-zip PPTX file | Catches `PackageNotFoundError` / `BadZipFile` and returns structured `InvalidPresentationError` detailing corrupted container. |
| 3 | `ppt_save` | Read-only destination file or locked file | Returns `FileAccessError` with actionable instruction to save to alternative location via `ppt_save_as`. |
| 4 | `ppt_revert` | Revert requested when no backup exists | Returns `BackupNotFoundError` with available session history and prevents data corruption. |
| 5 | `ppt_revert` | Revert with specific timestamp string | Replaces `working.pptx` with exact backup snapshot matching timestamp and returns confirmation. |
| 6 | `ppt_modify_shape` | Non-existent shape ID on slide | Returns `ShapeNotFound` with list of valid shape IDs and names currently present on the slide. |
| 7 | `ppt_modify_shape` | Negative or zero width/height | Returns `InvalidCoordinateError` with error message explaining dimensions must be positive floating point numbers. |
| 8 | `ppt_modify_text` | Shape has no text frame (e.g. pure image or connector) | Returns `InvalidTextFrameError` explaining shape type does not support text frames and suggests appropriate tool. |
| 9 | `ppt_modify_text` | Single run editing where run_index is out of bounds | Returns `IndexError` with total available run count and paragraph text snippet for context. |
| 10 | `ppt_validate_slide` | Background rectangle covering 100% of slide | Background rectangle is identified by z-order=0, 100% slide dimensions, and excluded from overlap collision warnings. |
| 11 | `ppt_validate_slide` | Grouped shapes or parent/child relationships | Inner shapes belonging to the same group container are evaluated relative to group bounds rather than triggering false sibling overlaps. |
| 12 | `ppt_validate_slide` | Shapes with 0.001-inch contact border (snapped edges) | Below minimum collision threshold (0.01 inches / 9144 EMUs), so not flagged as an overlap warning. |
| 13 | `ppt_render_slide` | PowerPoint COM unavailable and LibreOffice missing | Gracefully catches absence, returns structured `RendererUnavailableError` explaining neither COM nor LibreOffice was located. |
| 14 | `ppt_render_slide` | Simultaneous COM automation requests | Synchronized COM lock ensures serialized PowerPoint application access to prevent RPC server busy/collision errors. |
| 15 | `ppt_visual_diff` | Before and after images of differing pixel dimensions | Automatically normalizes/resizes images to common canvas with warning in response metadata before computing difference matrix. |
| 16 | `ppt_modify_ooxml` | Malformed XML snippet with invalid namespace or unclosed tag | Lxml validation catches XML syntax error, aborts modification without corrupting `working.pptx`, and returns XML parser error. |
| 17 | `ppt_delete_shape` | Attempt to delete the only slide layout placeholder required by master | Deletes shape from slide instance while preserving slide layout master definitions. |
| 18 | `mcp_config.json` | Existing user `mcp_config.json` with other MCP servers (e.g. `sqlite-helper`) | Config merge utility preserves existing server entries and seamlessly adds/updates only the `"powerpoint-mcp"` entry. |


# 6. Detailed Technical Specifications: R3 - Session & Safety Layer

## 6.1 Session & Working Copy Architecture

### Directory Hierarchy
```
<workspace_root>/
└── .ppt-agent/
    ├── config.json                     # Session manager configuration
    └── sessions/
        └── <session_id>/               # UUID4 session identifier (e.g. 3f8a1c92-...)
            ├── working.pptx            # Active modified working copy
            ├── original.pptx           # Copy of original source presentation
            ├── metadata.json           # Session status, paths, timestamps, history
            ├── backups/                # Chronological pre-mutation snapshots
            │   ├── presentation.backup-20260821-111530.pptx
            │   └── presentation.backup-20260821-112045.pptx
            ├── renders/                # Rendered PNG slide images
            │   ├── slide-01.png
            │   └── slide-02.png
            └── diffs/                  # Visual difference artifacts
                ├── diff-slide-01-v1-v2.png
                └── diff-slide-01-regions.json
```

### `metadata.json` Schema
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "SessionMetadata",
  "type": "object",
  "required": ["session_id", "source_path", "working_path", "created_at", "last_modified_at", "backups", "renders"],
  "properties": {
    "session_id": {"type": "string", "format": "uuid"},
    "source_path": {"type": "string"},
    "working_path": {"type": "string"},
    "created_at": {"type": "string", "format": "date-time"},
    "last_modified_at": {"type": "string", "format": "date-time"},
    "slide_count": {"type": "integer", "minimum": 0},
    "backups": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["timestamp", "backup_path", "operation"],
        "properties": {
          "timestamp": {"type": "string", "format": "date-time"},
          "backup_path": {"type": "string"},
          "operation": {"type": "string"},
          "details": {"type": "object"}
        }
      }
    },
    "renders": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["slide_number", "render_path", "renderer", "timestamp"],
        "properties": {
          "slide_number": {"type": "integer"},
          "render_path": {"type": "string"},
          "renderer": {"type": "string", "enum": ["powerpoint_com", "libreoffice", "mock"]},
          "timestamp": {"type": "string", "format": "date-time"}
        }
      }
    }
  }
}
```

### Session Lifecycle API (`SessionManager`)
1. **`create_session(source_path: str) -> Session`**:
   - Generates UUID4 session ID.
   - Creates directory tree under `.ppt-agent/sessions/<session_id>/`.
   - Copies `source_path` to `original.pptx` and `working.pptx`.
   - Writes initial `metadata.json`.
   - Sets active session pointer.
2. **`create_backup(operation: str, details: Optional[dict] = None) -> str`**:
   - Checks `PPT_BACKUP_ENABLED` (defaults to True).
   - Generates timestamp `YYYYMMDD-HHMMSS`.
   - Copies current `working.pptx` to `backups/presentation.backup-<timestamp>.pptx`.
   - Appends entry to `metadata.json`.
   - Returns absolute backup file path.
3. **`save(destination_path: Optional[str] = None) -> dict`**:
   - If `destination_path` is None, saves to original `source_path`.
   - Creates a timestamped backup of the target destination file if it exists before overwriting.
   - Copies `working.pptx` to destination.
   - Updates `metadata.json` with save record.
4. **`save_as(output_path: str, overwrite: bool = False) -> dict`**:
   - Resolves absolute path.
   - Creates parent directories if missing.
   - If destination exists and `overwrite=False`, raises `FileExistsError`.
   - If destination exists and `overwrite=True`, creates backup before overwriting.
   - Copies `working.pptx` to `output_path`.
5. **`revert(target: str = "original") -> dict`**:
   - If `target == "original"`, copies `original.pptx` over `working.pptx`.
   - If `target` is a backup timestamp or filename, locates backup in `backups/` and copies over `working.pptx`.
   - Records revert action in `metadata.json`.

---

## 6.2 Rule-Based Slide Validation (`ppt_validate_slide`)

### Validation Engine Rules & Thresholds

| Rule ID | Rule Name | Detection Algorithm | Threshold / Parameters | Diagnostic Output Format |
|---------|-----------|---------------------|------------------------|--------------------------|
| `VAL-01` | Overlap Detection | Axis-Aligned Bounding Box (AABB) intersection: `max(0, min(r1.right, r2.right) - max(r1.left, r2.left)) * max(0, min(r1.bottom, r2.bottom) - max(r1.top, r2.top))` | Min overlap area > 0.01 sq in (8.36e7 sq EMU). Ignore full-slide background rectangles (width >= 98% slide width and z-order=0). | `WARNING: Shape {id1} ('{name1}') overlaps Shape {id2} ('{name2}') by {depth_in:.2f} inches (area: {area_in:.2f} sq in).` |
| `VAL-02` | Boundary Clipping | Tests if `shape.left < 0`, `shape.top < 0`, `shape.left + shape.width > slide.width`, or `shape.top + shape.height > slide.height`. | Tolerance: > 0.05 inches protrusion. | `WARNING: Shape {id} ('{name}') extends {protrusion_in:.2f} inches beyond the {boundary} slide boundary.` |
| `VAL-03` | Off-Slide Placement | Tests if entire shape is outside slide canvas (`shape.right <= 0` or `shape.bottom <= 0` or `shape.left >= slide.width` or `shape.top >= slide.height`). | Exact canvas boundary. | `WARNING: Shape {id} ('{name}') is located completely off-slide at (x={x:.2f}, y={y:.2f}).` |
| `VAL-04` | Text Overflow Heuristic | Computes estimated text area: `num_chars * (font_size_pt / 72 * 0.55) * (font_size_pt / 72 * 1.25) * line_count`. Compares against inner frame area: `(width - left_margin - right_margin) * (height - top_margin - bottom_margin)`. | Estimated ratio > 1.15 (115% box capacity) and `word_wrap=True`. | `WARNING: Text in Shape {id} ('{name}') likely overflows text frame by estimated {overflow_pct}% (chars: {chars}, box: {w:.2f}x{h:.2f} in).` |
| `VAL-05` | Tiny Font Detection | Traverses all paragraphs and runs checking `font.size.pt`. | Threshold: font size < 8.0 pt (body/diagram) or < 6.0 pt (sub-footer). | `WARNING: Shape {id} ('{name}') contains suspiciously tiny text ({pt:.1f} pt): '{text_preview}'...` |
| `VAL-06` | Inconsistent Title Position | Compares title shape `(left, top, width, height)` against slide 1 title or presentation master title. | Delta threshold: `abs(delta_x) > 0.05 in` or `abs(delta_y) > 0.05 in`. | `WARNING: Title Shape {id} position (x={x:.2f}, y={y:.2f}) deviates from standard title position (x={std_x:.2f}, y={std_y:.2f}) by dx={dx:.2f}, dy={dy:.2f} in.` |
| `VAL-07` | Duplicate Objects | Detects shapes with identical `(shape_type, left, top, width, height)` and identical text or fill properties. | Exact coordinate match within 0.001 in. | `WARNING: Shape {id1} ('{name1}') and Shape {id2} ('{name2}') appear to be duplicate superimposed objects at (x={x:.2f}, y={y:.2f}).` |
| `VAL-08` | Extreme Rotation | Detects shape rotations outside standard angles. | `rotation % 90 != 0` and `rotation not in [45, 135, 225, 315]`. | `WARNING: Shape {id} ('{name}') has an irregular rotation of {rot:.1f} degrees.` |
| `VAL-09` | Unusually Large Images | Inspects image dimensions and raw byte sizes. | Image width > 1.5 * slide width or byte size > 25 MB. | `WARNING: Image Shape {id} ('{name}') has unusually large dimensions ({w:.2f}x{h:.2f} in) or byte size ({size_mb:.1f} MB).` |
| `VAL-10` | Inconsistent Margins | Compares text frame margins `(left_margin, top_margin, right_margin, bottom_margin)` across sibling shapes of the same semantic role. | Margin variance > 0.05 in. | `WARNING: Shape {id} has non-standard margins {margins} compared to sibling {role} shapes.` |

### Slide Validation Structured Return JSON Schema
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "SlideValidationReport",
  "type": "object",
  "required": ["slide_number", "is_valid", "warning_count", "error_count", "warnings", "metrics"],
  "properties": {
    "slide_number": {"type": "integer"},
    "is_valid": {"type": "boolean"},
    "warning_count": {"type": "integer"},
    "error_count": {"type": "integer"},
    "warnings": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["rule_id", "severity", "shape_ids", "message", "details"],
        "properties": {
          "rule_id": {"type": "string"},
          "severity": {"type": "string", "enum": ["info", "warning", "error"]},
          "shape_ids": {"type": "array", "items": {"type": "integer"}},
          "message": {"type": "string"},
          "details": {"type": "object"}
        }
      }
    },
    "metrics": {
      "type": "object",
      "properties": {
        "shape_count": {"type": "integer"},
        "text_shape_count": {"type": "integer"},
        "image_shape_count": {"type": "integer"},
        "slide_dimensions": {
          "type": "object",
          "properties": {
            "width_inches": {"type": "number"},
            "height_inches": {"type": "number"}
          }
        }
      }
    }
  }
}
```

# 7. Detailed Technical Specifications: R4 - Antigravity Integration & Tooling

## 7.1 FastMCP Server & 19 Core Tools

The server is built with FastMCP from the official MCP Python SDK:
```python
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    name="powerpoint-mcp",
    instructions="Deterministic PowerPoint (.pptx) inspection, editing, rendering, and validation MCP server for Antigravity."
)
```

### Complete 19 MCP Tools Catalog

#### 1. `ppt_open`
- **Description:** Open a PowerPoint presentation, initialize an isolated editing session with a working copy, and return session status and presentation overview.
- **Input Schema:**
  - `presentation_path` (str, required): Absolute or relative path to the `.pptx` file.
- **Return Type:**
  ```json
  {
    "session_id": "3f8a1c92-4b2a-4a21-9e8c-843810a9c84e",
    "working_path": "C:/.../.ppt-agent/sessions/3f8a1c92-.../working.pptx",
    "slide_count": 3,
    "dimensions": {"width_inches": 13.333, "height_inches": 7.5},
    "theme": "Office Theme",
    "titles": ["Company Overview", "Technical Architecture", "Issue Demonstration"]
  }
  ```
- **Error Behavior:** Raises `FileNotFoundError` if path does not exist, `InvalidPresentationError` if file is corrupt.

#### 2. `ppt_inspect_presentation`
- **Description:** Inspect high-level presentation metadata, slide count, dimensions, master layout names, and titles without modifying any state.
- **Input Schema:**
  - `presentation_path` (str, optional): Presentation path. If omitted, uses active session.
- **Return Type:**
  ```json
  {
    "slide_count": 3,
    "width_inches": 13.333,
    "height_inches": 7.5,
    "slide_width_emu": 12192000,
    "slide_height_emu": 6858000,
    "layouts": ["Title Slide", "Title and Content", "Section Header", "Two Content", "Blank"],
    "slides": [
      {"slide_number": 1, "title": "Company Overview", "shape_count": 6},
      {"slide_number": 2, "title": "Technical Architecture", "shape_count": 8},
      {"slide_number": 3, "title": "Issue Demonstration", "shape_count": 5}
    ]
  }
  ```

#### 3. `ppt_inspect_slide`
- **Description:** Inspect all shapes on a specific slide, returning geometry (in inches), semantic roles (title, subtitle, body, diagram, image, footer), text content, font styling, colors, and layout structure.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `presentation_path` (str, optional): Presentation path (defaults to active session).
- **Return Type:**
  ```json
  {
    "slide_number": 1,
    "shape_count": 6,
    "shapes": [
      {
        "shape_id": 2,
        "name": "Title 1",
        "role": "title",
        "shape_type": "TEXT_BOX",
        "x_inches": 1.0,
        "y_inches": 0.8,
        "width_inches": 11.333,
        "height_inches": 1.2,
        "rotation": 0.0,
        "z_order": 0,
        "text": "Company Overview",
        "font_family": "Calibri",
        "font_size_pt": 40.0,
        "bold": true,
        "italic": false,
        "color": "#1F497D",
        "alignment": "LEFT"
      }
    ]
  }
  ```

#### 4. `ppt_inspect_shape`
- **Description:** Get exhaustive details for a single shape on a slide, including all text frames, paragraphs, runs, margins, line styling, fill properties, and OOXML snippet.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, required): ID of the target shape.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** Detailed shape object with `paragraphs: List[ParagraphModel]`, `fill: FillModel`, `line: LineModel`, `xml_snippet: str`.

#### 5. `ppt_modify_shape`
- **Description:** Deterministically update a shape's coordinates, dimensions, rotation, z-order, or apply multi-shape alignment and distribution. Only provided properties are modified.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, required): Target shape ID.
  - `x` (float, optional): New X coordinate in inches.
  - `y` (float, optional): New Y coordinate in inches.
  - `width` (float, optional): New width in inches.
  - `height` (float, optional): New height in inches.
  - `rotation` (float, optional): Rotation angle in degrees (0-360).
  - `z_order` (str | int, optional): `"bring_to_front"`, `"send_to_back"`, `"bring_forward"`, `"send_backward"`, or absolute integer.
  - `align` (str, optional): `"left"`, `"center"`, `"right"`, `"top"`, `"middle"`, `"bottom"`.
  - `distribute` (str, optional): `"horizontal"`, `"vertical"`.
  - `target_shape_ids` (List[int], optional): Additional shape IDs for align/distribute operations.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:**
  ```json
  {
    "success": true,
    "shape_id": 2,
    "updated_properties": {"x_inches": 0.8, "y_inches": 0.8},
    "shape": {"shape_id": 2, "x_inches": 0.8, "y_inches": 0.8, "width_inches": 11.333, "height_inches": 1.2}
  }
  ```

#### 6. `ppt_modify_text`
- **Description:** Update text content and formatting in a text frame. Supports replacing full text or individual runs while strictly preserving surrounding rich-text styling.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, required): Target shape ID.
  - `text` (str, optional): New text string.
  - `font_family` (str, optional): Font name (e.g. "Calibri", "Aptos", "Arial").
  - `font_size` (float, optional): Font point size.
  - `bold` (bool, optional): Bold formatting.
  - `italic` (bool, optional): Italic formatting.
  - `underline` (bool, optional): Underline formatting.
  - `color` (str, optional): Hex RGB color string (e.g. "#1F497D").
  - `alignment` (str, optional): `"LEFT"`, `"CENTER"`, `"RIGHT"`, `"JUSTIFY"`.
  - `paragraph_spacing` (float, optional): Space after paragraph in points.
  - `line_spacing` (float, optional): Line spacing in points or lines.
  - `margins` (dict, optional): `{"left": float, "top": float, "right": float, "bottom": float}` in inches.
  - `paragraph_index` (int, optional): Specific 0-indexed paragraph to modify.
  - `run_index` (int, optional): Specific 0-indexed run to modify.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"success": true, "shape_id": 2, "text_summary": "Updated text preview...", "paragraph_count": 1}`.

#### 7. `ppt_copy_shape`
- **Description:** Clone an existing shape with all formatting, fills, lines, and text styles preserved onto the same or a different slide.
- **Input Schema:**
  - `slide_number` (int, required): Source slide number.
  - `shape_id` (int, required): ID of shape to copy.
  - `target_slide_number` (int, optional): Target slide number (defaults to same slide).
  - `x_offset` (float, optional): X offset in inches from original position (defaults to 0.2).
  - `y_offset` (float, optional): Y offset in inches from original position (defaults to 0.2).
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"success": true, "new_shape_id": 14, "target_slide": 1, "shape": {...}}`.

#### 8. `ppt_move_shape`
- **Description:** Move a shape by specifying absolute coordinates `(x, y)` or relative deltas `(dx, dy)` in inches.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, required): Target shape ID.
  - `dx` (float, optional): Delta X in inches.
  - `dy` (float, optional): Delta Y in inches.
  - `x` (float, optional): Absolute X in inches.
  - `y` (float, optional): Absolute Y in inches.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"success": true, "shape_id": 2, "x_inches": 0.8, "y_inches": 0.8}`.

#### 9. `ppt_resize_shape`
- **Description:** Resize a shape using absolute width/height or scaling factors with optional aspect ratio lock.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, required): Target shape ID.
  - `width` (float, optional): Absolute width in inches.
  - `height` (float, optional): Absolute height in inches.
  - `scale_width` (float, optional): Width multiplier (e.g. 1.2 for +20%).
  - `scale_height` (float, optional): Height multiplier.
  - `lock_aspect_ratio` (bool, optional): Maintain aspect ratio during scale.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"success": true, "shape_id": 2, "width_inches": 4.5, "height_inches": 3.0}`.

#### 10. `ppt_delete_shape`
- **Description:** Delete a shape cleanly from a slide.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, required): ID of shape to delete.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"success": true, "deleted_shape_id": 5, "remaining_shape_count": 4, "remaining_shape_ids": [2, 3, 4, 6]}`.

#### 11. `ppt_modify_ooxml`
- **Description:** Controlled low-level OOXML manipulation helper for features unsupported by python-pptx. Creates safety backup before mutation.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `shape_id` (int, optional): Target shape ID (or None for slide root).
  - `operation` (str, required): `"set_attribute"`, `"remove_attribute"`, `"insert_element"`, `"replace_element"`.
  - `xpath` (str, optional): XPath expression to target element.
  - `attributes` (dict, optional): Attribute dictionary to set/update.
  - `xml_fragment` (str, optional): Valid XML snippet to insert or replace.
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"success": true, "operation": "set_attribute", "affected_elements": 1, "xml_snippet": "<p:sp>..."}`.

#### 12. `ppt_validate_slide`
- **Description:** Run rule-based geometric and typographic validation on a slide, detecting overlaps, clipping, off-slide elements, tiny fonts, and text overflow.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `rules` (List[str], optional): Specific rule IDs to execute (defaults to all rules).
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** SlideValidationReport object (see Section 6.2 schema).

#### 13. `ppt_render_slide`
- **Description:** Render a single slide to high-resolution PNG using PowerPoint COM (Windows) or LibreOffice headless.
- **Input Schema:**
  - `slide_number` (int, required): 1-indexed slide number.
  - `output_dir` (str, optional): Directory to save PNG (defaults to `.ppt-agent/renders/<session_id>/`).
  - `renderer` (str, optional): `"auto"`, `"powerpoint_com"`, `"libreoffice"`.
  - `dpi` (int, optional): Image resolution DPI (defaults to 150).
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:**
  ```json
  {
    "image_path": "C:/.../.ppt-agent/renders/session-1/slide-01.png",
    "slide_number": 1,
    "renderer": "powerpoint_com",
    "width_px": 2000,
    "height_px": 1125
  }
  ```

#### 14. `ppt_render_presentation`
- **Description:** Render all slides in the presentation to PNG images.
- **Input Schema:**
  - `output_dir` (str, optional): Output directory path.
  - `renderer` (str, optional): Renderer choice (`"auto"`, `"powerpoint_com"`, `"libreoffice"`).
  - `dpi` (int, optional): DPI resolution (defaults to 150).
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"slide_count": 3, "rendered_slides": [{"slide_number": 1, "image_path": "..."}, ...], "renderer": "powerpoint_com"}`.

#### 15. `ppt_compare_slides`
- **Description:** Compare geometric, typographic, and semantic layout properties between two slides, matching corresponding shapes with confidence scores.
- **Input Schema:**
  - `slide_a` (int, required): First slide number (reference slide).
  - `slide_b` (int, required): Second slide number (target slide).
  - `match_shapes` (bool, optional): Perform semantic shape matching (defaults to True).
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:**
  ```json
  {
    "slide_a": 1,
    "slide_b": 2,
    "shape_count_a": 6,
    "shape_count_b": 6,
    "matches": [
      {
        "shape_a_id": 2,
        "shape_b_id": 2,
        "role": "title",
        "confidence": 0.96,
        "reasons": ["same semantic role: title", "same position (y=0.8 in)", "similar font size"],
        "geometric_differences": {"dx_inches": 0.0, "dy_inches": 0.0, "dwidth_inches": 0.0, "dheight_inches": 0.0},
        "style_differences": {"font_size_delta": 0.0, "color_match": true}
      }
    ]
  }
  ```

#### 16. `ppt_visual_diff`
- **Description:** Perform deterministic pixel-level image comparison between two slide renders, generating a diff image, changed bounding boxes, and similarity metrics.
- **Input Schema:**
  - `before_image` (str, required): Path to before PNG image.
  - `after_image` (str, required): Path to after PNG image.
  - `output_diff_path` (str, optional): Path to save diff image (defaults to `.ppt-agent/diffs/`).
  - `threshold` (float, optional): Pixel difference sensitivity threshold (0.0 - 1.0, defaults to 0.1).
- **Return Type:**
  ```json
  {
    "similarity_score": 0.9842,
    "pixel_difference_percentage": 1.58,
    "diff_image_path": "C:/.../.ppt-agent/diffs/diff_slide1.png",
    "changed_regions": [
      {"box_index": 1, "x_min": 120, "y_min": 95, "x_max": 450, "y_max": 230, "area_px": 44550}
    ]
  }
  ```

#### 17. `ppt_save`
- **Description:** Save session working copy back to the original presentation path, creating a timestamped backup before writing.
- **Input Schema:**
  - `presentation_path` (str, optional): Target path (defaults to session source path).
- **Return Type:** `{"saved_path": "C:/.../presentation.pptx", "backup_created": "C:/.../presentation.backup-20260821-113000.pptx", "success": true}`.

#### 18. `ppt_save_as`
- **Description:** Save session working copy to a specified new destination file path.
- **Input Schema:**
  - `output_path` (str, required): Destination `.pptx` file path.
  - `overwrite` (bool, optional): Whether to overwrite existing destination file (defaults to False).
  - `presentation_path` (str, optional): Source presentation path.
- **Return Type:** `{"saved_path": "C:/.../output/final.pptx", "backup_created": null, "success": true}`.

#### 19. `ppt_revert`
- **Description:** Discard current uncommitted edits and revert working copy to original state or a specified backup timestamp.
- **Input Schema:**
  - `target` (str, optional): `"original"` or backup timestamp string (defaults to `"original"`).
  - `presentation_path` (str, optional): Presentation path.
- **Return Type:** `{"reverted_to": "original", "working_path": "C:/.../working.pptx", "success": true}`.

---

## 7.2 Structured Error Handling Specification

All tool exceptions inherit from `PPTError` and return structured JSON dictionaries:
```python
class PPTError(Exception):
    def __init__(self, message: str, error_type: str, details: dict = None, available_options: list = None):
        super().__init__(message)
        self.message = message
        self.error_type = error_type
        self.details = details or {}
        self.available_options = available_options or []

    def to_dict(self):
        return {
            "success": False,
            "error_type": self.error_type,
            "message": self.message,
            "details": self.details,
            "available_options": self.available_options
        }
```

### Specific Error Types

| Exception Class | Error Type String | Typical Cause | Response Payload Content |
|-----------------|-------------------|---------------|--------------------------|
| `ShapeNotFoundError` | `"ShapeNotFound"` | Target `shape_id` does not exist on slide | `{"message": "Shape ID 17 does not exist on slide 2.", "details": {"slide_number": 2, "requested_id": 17}, "available_options": [2, 3, 4, 8]}` |
| `SlideNotFoundError` | `"SlideNotFound"` | `slide_number` < 1 or > `slide_count` | `{"message": "Slide 5 does not exist. Presentation has 3 slides.", "details": {"slide_number": 5, "total_slides": 3}, "available_options": [1, 2, 3]}` |
| `SessionNotFoundError`| `"SessionNotFound"`| No active session initialized and no path provided | `{"message": "No active editing session found. Please call ppt_open first.", "details": {}, "available_options": ["ppt_open"]}` |
| `InvalidCoordinateError`| `"InvalidCoordinate"`| Negative dimensions or invalid coordinate values | `{"message": "Width and height must be positive numbers.", "details": {"width": -2.0, "height": 3.0}}` |
| `InvalidTextFrameError` | `"InvalidTextFrame"` | Attempting text edit on shape without text support | `{"message": "Shape 4 (Picture) does not support text frames.", "details": {"shape_type": "PICTURE"}}` |
| `OOXMLValidationError` | `"OOXMLValidationError"` | Invalid XML fragment or malformed XPath | `{"message": "XML syntax error: unclosed tag <p:spPr>", "details": {"parser_error": "line 2 col 14"}}` |
| `RendererUnavailableError`| `"RendererUnavailable"`| Requested rendering backend not installed | `{"message": "PowerPoint COM and LibreOffice are not available on this host.", "details": {"tried": ["powerpoint_com", "libreoffice"]}}` |
| `BackupNotFoundError` | `"BackupNotFound"` | Revert target timestamp not found in backups directory | `{"message": "Backup timestamp 20260821-090000 not found.", "available_options": ["20260821-111530", "original"]}` |

---

## 7.3 MCP Resources

The MCP server exposes 3 standard URI resources:

1. **`ppt://current/presentation`**:
   - **MIME Type:** `application/json`
   - **Content:** Concise summary of presentation metadata, slide count, dimensions, master layouts, and slide title index.
2. **`ppt://current/slide/{slide_number}`**:
   - **MIME Type:** `application/json`
   - **Content:** Full structured shape tree of slide `{slide_number}`, with roles, coordinates, text snippets, and styling.
3. **`ppt://current/slide/{slide_number}/render`**:
   - **MIME Type:** `image/png`
   - **Content:** Binary PNG render of slide `{slide_number}` produced by the active renderer.

---

## 7.4 CLI Debugging Utilities

### 1. `scripts/inspect_pptx.py`
- **Purpose:** Standalone CLI to inspect presentation metadata, slides, or individual shapes without launching Antigravity.
- **Syntax:**
  ```powershell
  python scripts/inspect_pptx.py <presentation_path> [--slide <N>] [--shape <ID>] [--json] [--verbose]
  ```
- **Arguments:**
  - `presentation_path` (positional, required): Path to `.pptx` file.
  - `--slide <N>` (int, optional): Inspect specific slide number.
  - `--shape <ID>` (int, optional): Inspect specific shape ID on the specified slide.
  - `--json` (flag, optional): Output raw structured JSON instead of human-readable colored ASCII tree.
  - `--verbose` (flag, optional): Include raw XML snippets and advanced properties.
- **Exit Codes:**
  - `0`: Success.
  - `1`: File not found or invalid argument.
  - `2`: Presentation parse error.

### 2. `scripts/render_pptx.py`
- **Purpose:** Standalone CLI to render presentation slides to PNG images for debugging.
- **Syntax:**
  ```powershell
  python scripts/render_pptx.py <presentation_path> [--slide <N>] [--output <DIR>] [--renderer <auto|com|libreoffice>] [--dpi <N>]
  ```
- **Arguments:**
  - `presentation_path` (positional, required): Path to `.pptx` file.
  - `--slide <N>` (int, optional): Render only specific slide (defaults to all slides).
  - `--output <DIR>` (str, optional): Output directory (defaults to `./renders`).
  - `--renderer <choice>` (str, optional): Renderer selection (`auto`, `com`, `libreoffice`).
  - `--dpi <N>` (int, optional): Render resolution DPI (defaults to 150).
- **Exit Codes:**
  - `0`: Success (prints rendered file paths to stdout).
  - `1`: Renderer error or invalid argument.

---

## 7.5 Antigravity Configuration (`.agents/mcp_config.json`)

### Exact Workspace Configuration
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "command": "uv",
      "args": [
        "run",
        "--with-editable",
        ".",
        "python",
        "-m",
        "powerpoint_mcp.server"
      ],
      "env": {
        "PPT_RENDERER": "auto",
        "PPT_WORKSPACE_DIR": ".ppt-agent",
        "PPT_BACKUP_ENABLED": "true",
        "PPT_DEFAULT_OUTPUT_DIR": "./output"
      }
    }
  }
}
```

### Safe Merging Logic
When deploying `.agents/mcp_config.json`:
1. Check if `.agents/mcp_config.json` already exists.
2. If it exists, read and parse JSON into a Python dictionary.
3. If `"mcpServers"` key is missing, initialize `"mcpServers": {}`.
4. Inject/update `"powerpoint-mcp"` entry under `"mcpServers"`.
5. Preserve all other existing server entries (e.g. `sqlite-helper`, `github-mcp`).
6. Serialize back to `.agents/mcp_config.json` with 2-space indentation.

## 7.6 Antigravity Skill (`.agents/skills/powerpoint-editor/SKILL.md`)

### File Path: `.agents/skills/powerpoint-editor/SKILL.md`

### YAML Frontmatter & Content Specification
```markdown
---
name: powerpoint-editor
description: >-
  Expert conversational PowerPoint (.pptx) editor. Teaches the agent to make deterministic,
  minimal-diff edits, inspect shape geometry, maintain style fidelity, perform rule-based validation,
  and verify visual results using rendering loops.
---

# PowerPoint Editor Skill

You are a precision conversational PowerPoint editor. Your objective is to modify existing presentations with surgical accuracy, preserving existing layouts, themes, relationships, and formatting.

## 15 Immutable PowerPoint Editing Rules

1. **Always inspect before editing**: Never modify a slide or shape without first calling `ppt_inspect_slide` or `ppt_inspect_shape`.
2. **Identify objects semantically**: Reference shapes by their semantic roles (`title`, `subtitle`, `body`, `diagram`, `image`, `footer`) and shape IDs.
3. **Make the smallest possible change**: Apply minimal-diff edits. Never recreate or replace shapes when modifying properties suffices.
4. **Preserve existing styles**: When updating text, preserve font family, font size, colors, and paragraph spacing unless explicitly instructed.
5. **Never recreate an object when it can be modified**: Modify existing coordinates and text frames in-place.
6. **Never rebuild an entire slide**: Only adjust specific target elements requested by the user.
7. **Render after visual changes**: Always call `ppt_render_slide` after modifying coordinates, dimensions, alignments, or text styling.
8. **Inspect the rendered result**: Review visual output or run `ppt_visual_diff` / `ppt_validate_slide` to verify changes.
9. **Correct if necessary**: If validation detects overlaps or misalignments, apply corrective adjustments immediately.
10. **Save only after verification**: Call `ppt_save` or `ppt_save_as` only after verifying visual and geometric correctness.
11. **Prefer exact geometric operations**: Use exact decimal inch coordinates or alignment tools rather than guessing coordinates.
12. **Inspect reference slides first**: When asked to "make slide A look like slide B", call `ppt_compare_slides` or inspect both slides before modifying slide A.
13. **Preserve target content during style matching**: Copy only layout, geometry, and styling from the reference slide; keep the target slide's text and assets intact.
14. **Do not alter unrelated slides**: Confine modifications strictly to the target slide(s).
15. **Inspect when uncertain**: If user instructions are ambiguous (e.g. "move the blue box"), inspect the slide to disambiguate shape IDs before modifying.

## Workflow Decision Trees

### 1. Text Modification Workflow
```
User asks to change text / typography
    ↓
ppt_inspect_slide (identify shape ID and current text/font)
    ↓
ppt_modify_text (update text/styling, preserving surrounding runs)
    ↓
ppt_render_slide (if font size or layout might shift)
    ↓
ppt_validate_slide (check for text overflow)
    ↓
Report changes to user
```

### 2. Geometry Modification Workflow
```
User asks to move, resize, align, or distribute shapes
    ↓
ppt_inspect_slide (get current coordinates and dimensions)
    ↓
Calculate exact target coordinates in inches
    ↓
ppt_modify_shape / ppt_move_shape / ppt_resize_shape
    ↓
ppt_validate_slide (check for overlaps and boundary clipping)
    ↓
ppt_render_slide (render affected slide)
    ↓
Verify visual integrity & report to user
```

### 3. Reference Slide Layout Matching Workflow
```
User asks: "Make slide A look like slide B"
    ↓
ppt_inspect_slide(slide_A) AND ppt_inspect_slide(slide_B)
    ↓
ppt_compare_slides(slide_B, slide_A) (match semantic shapes)
    ↓
Determine geometry & styling deltas
    ↓
Apply minimal modifications to slide A shapes (preserve slide A content!)
    ↓
ppt_render_slide(slide_A)
    ↓
ppt_validate_slide(slide_A)
    ↓
Verify visual match against slide B
```

### 4. Tool Call Optimization & Batching
- Do NOT make repetitive `ppt_inspect_shape` calls if `ppt_inspect_slide` already contains the required properties.
- Use `align` and `distribute` parameters in `ppt_modify_shape` for multi-shape adjustments in a single call.
```

---

# 8. Synthetic 3-Slide Test Presentation Specification

### File: `tests/fixtures/sample_presentation.pptx`
Widescreen 16:9 format (`width = 13.333 inches / 12192000 EMU`, `height = 7.5 inches / 6858000 EMU`).

### Slide 1: Company Overview
- **Title**: `Title 1` at `(x=1.0", y=0.8", w=11.333", h=1.0")`, Text: `"Company Overview"`, Font: Calibri 40pt Bold, Color: `#1F497D`. Role: `title`.
- **Subtitle**: `Subtitle 2` at `(x=1.0", y=1.9", w=11.333", h=0.6")`, Text: `"Q3 Performance & Future Strategic Roadmap"`, Font: Calibri 20pt Italic, Color: `#595959`. Role: `subtitle`.
- **Feature Box 1**: `Box 1` at `(x=1.0", y=2.8", w=3.4", h=3.2")`, Fill: `#E7EEF8`, Border: `#2E75B6`, Text: `"Cloud Growth

- Revenue up 42% YoY
- 150+ Enterprise Clients
- 99.99% SLA Met"`, Font: Calibri 14pt. Role: `body`.
- **Feature Box 2**: `Box 2` at `(x=4.966", y=2.8", w=3.4", h=3.2")`, Fill: `#E7EEF8`, Border: `#2E75B6`, Text: `"AI Capabilities

- Real-time Inference
- Autonomous Agents
- Multi-modal Pipelines"`, Font: Calibri 14pt. Role: `body`.
- **Feature Box 3**: `Box 3` at `(x=8.933", y=2.8", w=3.4", h=3.2")`, Fill: `#E7EEF8`, Border: `#2E75B6`, Text: `"Global Expansion

- 4 New Regional Hubs
- Localized Compliance
- 24/7 Operations"`, Font: Calibri 14pt. Role: `body`.
- **Image Placeholder**: `Image 1` at `(x=1.0", y=6.2", w=11.333", h=0.7")`, Fill: `#F2F2F2`, Text: `"Placeholder: Partner Ecosystem Badges"`. Role: `image`.

### Slide 2: Technical Architecture
- **Title**: `Title 1` at `(x=1.0", y=0.8", w=11.333", h=1.0")`, Text: `"Technical Architecture"`, Font: Calibri 40pt Bold, Color: `#1F497D`. Role: `title`.
- **Left Column (Body Bullet List)**: `Left Col` at `(x=1.0", y=2.0", w=5.4", h=4.5")`, Fill: None, Text: `"Core Components:
• Antigravity Agent Controller
• FastMCP Stdio Protocol Bridge
• Deterministic Geometry Engine
• PPTX/OOXML Mutation Layer
• COM / LibreOffice Visual Renderer"`, Font: Calibri 16pt. Role: `body`.
- **Right Column (Multi-Shape Diagram)**:
  - Diagram Container: `Diag Container` at `(x=6.9", y=2.0", w=5.4", h=4.5")`, Fill: `#FAFAFA`, Border: `#D9D9D9`.
  - Node A (Agent): `Node A` at `(x=7.4", y=2.4", w=4.4", h=0.9")`, Fill: `#4F81BD`, Text: `"Antigravity CLI Agent"`, Font: Calibri 14pt White Bold. Role: `diagram`.
  - Node B (MCP Server): `Node B` at `(x=7.4", y=3.7", w=4.4", h=0.9")`, Fill: `#9BBB59`, Text: `"PowerPoint MCP Server"`, Font: Calibri 14pt White Bold. Role: `diagram`.
  - Node C (Engine): `Node C` at `(x=7.4", y=5.0", w=4.4", h=0.9")`, Fill: `#8064A2`, Text: `"PowerPoint COM / python-pptx"`, Font: Calibri 14pt White Bold. Role: `diagram`.
- **Footer**: `Footer 1` at `(x=1.0", y=6.8", w=11.333", h=0.4")`, Text: `"Confidential — Internal Use Only"`, Font: Calibri 10pt Italic, Color: `#7F7F7F`. Role: `footer`.

### Slide 3: Issue Demonstration (Deliberate Validation Flaws)
- **Title**: `Title 1` at `(x=1.0", y=0.8", w=11.333", h=1.0")`, Text: `"Issue Demonstration"`, Font: Calibri 40pt Bold. Role: `title`.
- **Overlapping Shape A**: `Overlap A` at `(x=2.0", y=2.5", w=4.0", h=2.5")`, Fill: `#FFC000`, Text: `"Primary Module"`. Role: `body`.
- **Overlapping Shape B (Overlaps A by 0.35 inches)**: `Overlap B` at `(x=4.5", y=3.0", w=4.0", h=2.5")`, Fill: `#ED7D31`, Text: `"Secondary Module (Overlapping)"`. Role: `body`.
- **Off-Slide Shape (Extends 0.4 inches beyond right border)**: `Protruding Box` at `(x=11.5", y=2.5", w=2.233", h=2.5")` -> `x + w = 13.733" > 13.333"`. Text: `"Clipped Box"`. Role: `body`.
- **Tiny Font Box**: `Tiny Text Box` at `(x=2.0", y=5.5", w=4.0", h=1.0")`, Text: `"Unreadable fine print details"`, Font: Calibri 5.0 pt. Role: `body`.

---

# 9. Test Suite Specifications

### Test Structure & Coverage
```
tests/
├── fixtures/
│   ├── generate_test_presentation.py   # Script creating sample_presentation.pptx
│   └── sample_presentation.pptx        # Generated 3-slide test fixture
├── test_session.py                     # Working copy, backup lifecycle, revert, save_as
├── test_validation.py                  # Overlap, clipping, overflow, tiny font rules
├── test_mcp.py                         # In-memory FastMCP client testing all 19 tools & resources
├── test_inspection.py                  # Presentation, slide, shape metadata, roles
├── test_geometry.py                    # Move, resize, alignment, distribution
├── test_editing.py                     # Copy, text formatting, deletion, z-order
└── test_rendering.py                   # COM / LibreOffice detection, image diffing
```

### Key Test Case Specifications

#### `tests/test_session.py`
- `test_session_lifecycle`: Verify `create_session` provisions `.ppt-agent/sessions/<id>/` with `working.pptx` and `metadata.json`.
- `test_backup_generation`: Verify modifying working copy generates `presentation.backup-YYYYMMDD-HHMMSS.pptx`.
- `test_save_as_non_destructive`: Verify `ppt_save_as` writes new `.pptx` file while leaving original source intact.
- `test_revert_to_original`: Verify `ppt_revert` rolls back working copy to initial state.
- `test_revert_to_timestamp`: Verify `ppt_revert(target=timestamp)` restores exact historical snapshot.

#### `tests/test_validation.py`
- `test_validate_slide_clean`: Slide 1 validation returns `is_valid=True`, `warning_count=0`.
- `test_validate_slide_overlap_detection`: Slide 3 validation detects overlap between `Overlap A` and `Overlap B` with accurate depth/area metrics.
- `test_validate_slide_boundary_clipping`: Slide 3 validation detects `Protruding Box` extending 0.40 inches past right boundary.
- `test_validate_slide_tiny_font`: Slide 3 validation detects `Tiny Text Box` (5.0 pt text).

#### `tests/test_mcp.py` (In-Memory FastMCP Client)
- `test_mcp_tool_discovery`: In-memory client queries server and discovers all 19 tools with expected schemas.
- `test_mcp_resource_discovery`: In-memory client discovers 3 URI resources.
- `test_mcp_inspect_presentation_tool`: Invocations of `ppt_inspect_presentation` return structured dictionary.
- `test_mcp_modify_shape_tool`: Modifies shape coordinates and verifies return model.
- `test_mcp_modify_text_rich_preservation`: Modifies single run and verifies adjacent font styling is preserved.
- `test_mcp_error_handling_shape_not_found`: Passing non-existent shape ID returns structured `ShapeNotFound` JSON payload with valid shape list.

---

# 10. `README.md` Specification

The project `README.md` must contain the following 12 complete sections:
1. **Title & Overview**: Description of the PowerPoint MCP Server for Antigravity, core philosophy (minimal-diff, deterministic, visual verification).
2. **Architecture**: ASCII diagram illustrating Antigravity CLI -> Stdio MCP -> FastMCP -> python-pptx / COM / Renderer / Validator.
3. **Prerequisites & Requirements**: Python 3.10+, `uv`, Windows OS with optional Microsoft PowerPoint Office 16 or LibreOffice.
4. **Installation & Setup**: Commands using `uv sync` and virtual environment setup.
5. **Running the MCP Server**: Local stdio invocation command (`uv run python -m powerpoint_mcp.server`).
6. **Antigravity Workspace Configuration**: Explaining `.agents/mcp_config.json` and automatic discovery.
7. **Antigravity Skill Installation**: Explaining `.agents/skills/powerpoint-editor/SKILL.md` and progressive disclosure.
8. **Supported MCP Tools Reference**: Complete table of all 19 tools, inputs, outputs, and descriptions.
9. **Supported MCP Resources**: `ppt://current/presentation`, `ppt://current/slide/{N}`, `ppt://current/slide/{N}/render`.
10. **Standalone CLI Debugging Tools**: Usage of `scripts/inspect_pptx.py` and `scripts/render_pptx.py`.
11. **Example Conversational Commands**: 10+ real-world prompts (inspect, move, equalize, match slide, validate, render, save).
12. **Troubleshooting & FAQ**: COM automation cleanup, headless fallback, font issues, permission issues.
