# PowerPoint MCP Server for Antigravity

A production-quality local Model Context Protocol (MCP) server in Python enabling conversational, deterministic PowerPoint (`.pptx`) inspection, surgical editing, semantic component authoring, cross-slide harmonization, high-resolution rendering, visual diffing, and rule-based validation for Google Antigravity.

---

## 1. Overview & Philosophy

Modern AI slide editing often destroys fine-grained formatting, breaks font hierarchies, and resets complex layouts because LLMs attempt to regenerate whole presentations from scratch.

The **PowerPoint MCP Server** implements a deterministic, minimal-diff editing philosophy:
- **Inspect Before Mutating**: Semantic roles (`title`, `card`, `card_title`, `metric`, `badge`, `stepper`, `table`, `image`, `footer`) and exact EMU/inch bounding boxes are extracted directly from OpenXML trees.
- **Strict PATCH Semantics**: Shape modifications only mutate explicitly provided fields. Omitted coordinates or dimensions are never converted to zero or overwritten.
- **Semantic Components**: Native primitives for multi-step process flows (`ppt_create_stepper`, `ppt_update_stepper`), structured card lists (`ppt_create_structured_card_list`), and composite component manipulation (`ppt_move_component`, `ppt_resize_component`).
- **Cross-Slide Consistency**: Harmonize slide sequences using `ppt_sync_slide_chrome`, `ppt_sync_component`, and `ppt_sync_layout` without overwriting target slide content.
- **First-Class Media & Tables**: Native image insertion (`ppt_add_picture` with automatic aspect ratio calculation), image replacement (`ppt_replace_picture`), and cell-level table inspection, styling, merging, geometry, and batch editing.
- **Run-Level Style & Bullet Preservation**: Text edits preserve font faces, point sizes, colors, margins, bullet styles, indent levels, and multi-run formatting.
- **Non-Destructive Working Copies**: Modifications occur in isolated session workspaces under `.ppt-agent/sessions/<session_id>/working.pptx` with automatic timestamped backups (`presentation.backup-YYYYMMDD-HHMMSS.pptx`).
- **Visual Verification Loops**: Integrated Windows PowerPoint COM automation with process-wide file-based locking and headless LibreOffice rendering generate PNG slide images and pixel difference heatmaps to visually verify layout changes.

---

## 2. Architecture

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                 Antigravity CLI Agent                                  │
│                   (Skill: .agents/skills/powerpoint-editor/SKILL.md)                   │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │ JSON-RPC 2.0 via Stdio
┌───────────────────────────────────────────▼────────────────────────────────────────────┐
│                         PowerPoint MCP Server (server.py)                              │
│                    53 FastMCP / MCPServer Tools  &  3 MCP Resources                    │
└───────┬───────────────────────┬───────────────────────┬────────────────────────┬───────┘
        │                       │                       │                        │
┌───────▼──────────────┐ ┌──────▼───────────────┐ ┌─────▼────────────────┐ ┌─────▼───────┐
│ Inspection & Semantic│ │ Core & Layout Editing│ │ Media & Tables       │ │ Rendering   │
│ - Structure & Roles  │ │ - Move, Resize, Copy │ │ - Native Pictures    │ │ - PPT COM   │
│ - Container Inference│ │ - Align & Distribute │ │ - Image Replacement  │ │ - COM Lock  │
│ - Cross-Slide Diff   │ │ - Equalize & Space   │ │ - Table Cell Edits   │ │ - LibreOff. │
│ - Rule Validation    │ │ - Atomic Containers  │ │ - Table Geometry     │ │ - Visual    │
│ - Table Validation   │ │ - Relative Typo      │ │ - Cell Merge & Style │ │   Diffing   │
└──────────────────────┘ └──────────────────────┘ └──────────────────────┘ └─────────────┘
                                            │
┌───────────────────────────────────────────▼────────────────────────────────────────────┐
│                             Session & Safety Layer                                     │
│           .ppt-agent/sessions/<session_id>/working.pptx + timestamped backups          │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 3. Prerequisites & Requirements

- **Operating System**: Windows 10/11 (recommended for native PowerPoint COM automation), macOS, or Linux.
- **Python**: Python 3.10, 3.11, 3.12, or 3.13+.
- **Package Manager**: `uv` (recommended) or `pip`.
- **Presentation Software (Optional but recommended for rendering)**:
  - Microsoft PowerPoint 2016 / 2019 / 2021 / Office 365 (Windows).
  - LibreOffice (`soffice` CLI in system PATH).

---

## 4. Installation & Setup

Clone the repository and install dependencies using `uv`:

```bash
# Clone the repository
git clone https://github.com/example/powerpoint-mcp.git
cd powerpoint-mcp

# Create and sync virtual environment with uv
uv sync
```

Alternatively, with standard `pip`:

```powershell
python -m venv .venv
# On Windows PowerShell:
.\.venv\Scripts\Activate.ps1
# On macOS/Linux:
source .venv/bin/activate

pip install -e .
```

---

## 5. Running the MCP Server

To run the MCP server over standard input/output (stdio):

```bash
uv run python -m powerpoint_mcp.server
```

Or via direct script entrypoint:

```bash
uv run powerpoint-mcp
```

---

## 6. Antigravity Workspace Configuration

Configure your workspace `.agents/mcp_config.json` to register the server:

```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "command": "uv",
      "args": [
        "run",
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

To sync server schemas and the editing skill into the user's local Antigravity environment:

```powershell
python scripts/sync_mcp.py
```

---

## 7. Antigravity Skill Installation

The skill definition is located at `.agents/skills/powerpoint-editor/SKILL.md`. Antigravity automatically discovers skills inside `.agents/skills/`.

The skill encodes:
- The **22 Immutable PowerPoint Editing Rules** (covering session isolation, PATCH semantics, semantic components, layout primitives, media, tables, and validation).
- Structured decision trees for cross-slide flow harmonization, relative typography scaling, container manipulation, and card lists.
- Tool call optimization and batching rules.

---

## 8. Supported MCP Tools Reference (53 Core Tools)

### Session & Versioning Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_open` | `presentation_path: str` | Open presentation and initialize isolated session workspace with working copy. |
| `ppt_save` | `presentation_path: Optional[str]` | Save working copy back to original path with automatic timestamped backup. |
| `ppt_save_as` | `output_path: str, overwrite: bool, presentation_path` | Save working copy to a new destination path. |
| `ppt_revert` | `target: str, presentation_path` | Revert working copy to original presentation or a specified backup timestamp. |

### Inspection & Semantic Analysis Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_inspect_presentation` | `presentation_path: Optional[str]` | Inspect presentation metadata, slide count, dimensions, layouts, and titles. |
| `ppt_inspect_slide` | `slide_number: int, presentation_path: Optional[str]` | Inspect slide shapes, EMU/inch coords, semantic roles, text, and styles. |
| `ppt_inspect_text` | `slide_number: int, include_geometry, include_style, presentation_path` | Filtered text inspection with font hierarchy, alignment, and overflow warnings. |
| `ppt_inspect_shape` | `slide_number: int, shape_id: int, presentation_path: Optional[str]` | Deep inspection of single shape: runs, fonts, colors, margins, line, fill, XML. |
| `ppt_inspect_components` | `slide_number: int, detail: str, presentation_path: Optional[str]` | Inspect semantic visual components (headers, footers, steppers, cards, containers). |
| `ppt_inspect_table` | `slide_number: int, table_shape_id: Optional[int], detail: str, presentation_path` | Inspect table grid, column widths, row heights, bbox, and cell-level formatting. |
| `ppt_analyze_slide_structure` | `slide_number: int, presentation_path: Optional[str]` | Analyze semantic hierarchy, roles (title, card, metric, badge), and container trees. |
| `ppt_analyze_containers` | `slide_number: int, presentation_path: Optional[str]` | Identify logical card containers and nested child shapes with bounding boxes. |
| `ppt_compare_slides` | `slide_a: int, slide_b: int, reference_slide, target_slides, presentation_path` | Compare geometry, layout, typography, and semantic matches across slides. |
| `ppt_validate_slide` | `slide_number: int, rules: Optional[List[str]], detail: str, presentation_path` | Rule-based validation for overlaps (VAL-01), clipping (VAL-02), tiny fonts (VAL-04), overflow (VAL-03), and tables (TABLE-01/02). |
| `ppt_validate_slides` | `slide_numbers: Optional[List[int]], detail: bool, rules, presentation_path` | Multi-slide batch validation returning compact error/warning counts or full findings. |

### Core Shape & Text Editing Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_modify_shape` | `slide_number, shape_id, x, y, width, height, rotation, z_order, dx, dy, dwidth, dheight, drotation, align, distribute, presentation_path` | Modify coordinates, dimensions, rotation, or z-order with strict PATCH semantics. |
| `ppt_modify_text` | `slide_number, shape_id, text, font_family, font_size, font_size_delta, font_size_scale, bold, italic, color, alignment, presentation_path` | Modify text with run-level style and paragraph bullet preservation. |
| `ppt_batch_modify_text` | `slide_number, operations: List[Dict], presentation_path` | Batch text updates across multiple shapes in a single transaction. |
| `ppt_batch_modify_shapes` | `slide_number, operations: List[Dict], presentation_path` | Batch shape coordinate and dimension updates in a single transaction. |
| `ppt_copy_shape` | `slide_number, shape_id, target_slide_number, x_offset, y_offset, presentation_path` | Clone shape with formatting and relationships onto same or target slide. |
| `ppt_move_shape` | `slide_number, shape_id, dx, dy, x, y, presentation_path` | Move shape by absolute coordinates or relative offsets in inches. |
| `ppt_resize_shape` | `slide_number, shape_id, width, height, scale_width, scale_height, lock_aspect_ratio, presentation_path` | Resize shape with absolute dimensions or scaling multipliers. |
| `ppt_delete_shape` | `slide_number, shape_id, presentation_path` | Remove shape cleanly from slide shape tree. |
| `ppt_modify_ooxml` | `slide_number, shape_id, operation, xpath, xml_fragment, transparency_percent, gradient_start, gradient_end, presentation_path` | Safe low-level DrawingML manipulation for transparency, gradients, and shadows. |

### High-Level Layout & Container Primitives
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_align_shapes` | `slide_number, shape_ids: List[int], alignment: str, presentation_path` | Align shapes (`left`, `center`, `right`, `top`, `middle`, `bottom`). |
| `ppt_distribute_shapes` | `slide_number, shape_ids: List[int], direction: str, mode: str, presentation_path` | Distribute shapes horizontally or vertically with equal gaps or equal centers. |
| `ppt_space_shapes` | `slide_number, shape_ids: List[int], gap_inches: float, direction: str, presentation_path` | Set exact, uniform spacing between consecutive shapes in inches. |
| `ppt_equalize_sizes` | `slide_number, shape_ids: List[int], dimension: str, target: str, presentation_path` | Equalize width, height, or both across shapes based on max, min, or average. |
| `ppt_move_container` | `slide_number, container_id, dx, dy, x, y, presentation_path` | Atomically move card container and all contained child elements together. |
| `ppt_resize_container` | `slide_number, container_id, width, height, scale_width, scale_height, presentation_path` | Atomically resize card and proportionally scale child positions and dimensions. |
| `ppt_reflow_container` | `slide_number, container_id, padding_top, padding_bottom, gap_inches, presentation_path` | Vertically reflow and restack child elements inside a container with clean padding. |
| `ppt_scale_slide_typography` | `slide_number, scale_factor, min_font_size, max_font_size, presentation_path` | Scale all text elements across a slide proportionally while preserving hierarchy. |
| `ppt_apply_style` | `slide_number, shape_id, preset_name, source_shape_id, presentation_path` | Apply role-based style presets (`card_default`, `badge_success`, `title_hero`) or transfer styles. |
| `ppt_create_flow_diagram` | `slide_number, steps: List[Dict], left, top, width, height, orientation, presentation_path` | Generate multi-step process flow diagram with cards, titles, bodies, and connectors. |

### Semantic Component & Harmonization Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_create_stepper` | `slide_number, steps: List[Dict], left, top, width, height, style, presentation_path` | Generate standardized breadcrumb stepper component with active/completed/future states. |
| `ppt_update_stepper` | `slide_number, active_step, completed_steps, presentation_path` | Update stepper progression across a slide sequence without rebuilding shapes. |
| `ppt_sync_component` | `reference_slide, target_slides: List[int], component_type: str, match_by, presentation_path` | Synchronize visual style, position, or content of components across slides. |
| `ppt_sync_slide_chrome` | `reference_slide, target_slides: List[int], include_header, include_footer, include_stepper, presentation_path` | Synchronize slide chrome (header title, category banner, footer, stepper) across sequence. |
| `ppt_sync_layout` | `reference_slide, target_slides: List[int], component: str, preserve_content: bool, presentation_path` | Transfer structural layout and positioning to target slides while keeping content intact. |
| `ppt_create_structured_card_list` | `slide_number, cards: List[Dict], left, top, width, height, columns, presentation_path` | Generate grid of structured cards with badge, title, metric, description, and list items. |
| `ppt_move_component` | `slide_number, component_id: str, dx, dy, x, y, presentation_path` | Move composite component (e.g. `card_1`, `stepper`, `content_area`) atomically. |
| `ppt_resize_component` | `slide_number, component_id: str, width, height, scale_width, scale_height, reflow_children, presentation_path` | Resize composite component atomically while proportionally adjusting children. |

### Images & Media Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_add_picture` | `slide_number, image_path, left, top, width, height, preserve_aspect_ratio, presentation_path` | Insert PNG, JPEG, BMP image with automatic aspect ratio calculation. |
| `ppt_replace_picture` | `slide_number, shape_id, image_path, preserve_geometry, presentation_path` | Replace picture or rectangle placeholder in-place, preserving bounds and rotation. |

### Table Operations Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_batch_modify_table_cells` | `slide_number, table_shape_id, mutations: List[Dict], presentation_path` | Atomic multi-cell text and formatting updates (fonts, bold, fill, alignment). |
| `ppt_set_table_geometry` | `slide_number, table_shape_id, left, top, width, height, column_widths, row_heights, presentation_path` | Modify table bounds, individual column widths, or row heights with PATCH semantics. |
| `ppt_style_table` | `slide_number, table_shape_id, range: str, style: Dict, presentation_path` | Apply formatting (fill, fonts, borders, alignment, margins) across table or cell range. |
| `ppt_merge_table_cells` | `slide_number, table_shape_id, start_row, start_column, end_row, end_column, presentation_path` | Merge rectangular block of table cells with range validation. |
| `ppt_batch_modify_tables` | `operations: List[Dict], presentation_path` | Multi-slide, multi-table atomic transaction executing cell edits, geometry, styles, and merges. |

### Rendering & Verification Tools
| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_render_slide` | `slide_number, output_dir, output_path, renderer, dpi, presentation_path` | Render single slide to high-resolution PNG image with timing metadata. |
| `ppt_render_slides` | `slide_numbers: List[int], output_dir, renderer, dpi, presentation_path` | Batch render slides in a single PowerPoint COM session to minimize startup overhead. |
| `ppt_render_presentation` | `output_dir, renderer, dpi, presentation_path` | Render all slides in presentation to PNG images. |
| `ppt_visual_diff` | `before_image, after_image, output_diff_path, threshold, presentation_path` | Generate pixel-level diff heatmap, changed bounding boxes, and similarity score. |

---

## 9. Supported MCP Resources

| URI Resource | MIME Type | Description |
|---|---|---|
| `ppt://current/presentation` | `application/json` | JSON summary of presentation metadata, dimensions, layouts, and slide titles. |
| `ppt://current/slide/{slide_number}` | `application/json` | JSON structured shape tree with roles, coordinates, and typography for slide `{slide_number}`. |
| `ppt://current/slide/{slide_number}/render` | `image/png` | High-resolution PNG binary render of slide `{slide_number}`. |

---

## 10. Standalone CLI Debugging Tools

### Inspect Presentation or Slide
```powershell
# Inspect entire presentation metadata
python scripts/inspect_pptx.py presentation.pptx

# Inspect specific slide as ASCII tree
python scripts/inspect_pptx.py presentation.pptx --slide 1

# Inspect single shape in JSON format
python scripts/inspect_pptx.py presentation.pptx --slide 1 --shape 2 --json
```

### Render Slide to PNG
```powershell
# Render all slides to ./renders/
python scripts/render_pptx.py presentation.pptx --output ./renders

# Render single slide at 300 DPI using PowerPoint COM
python scripts/render_pptx.py presentation.pptx --slide 1 --dpi 300 --renderer powerpoint
```

### Sync MCP Schemas & Skill
```powershell
python scripts/sync_mcp.py
```

---

## 11. Example Conversational Commands

Here are typical natural-language prompts handled by Antigravity using the PowerPoint MCP Server:

1. *"Open deck.pptx and inspect the tables on slide 3."*
2. *"Populate the table on slide 3 with this CSV content, make the header row bold with dark blue fill, and adjust column widths so text fits."*
3. *"Insert architecture_diagram.png onto slide 4 at position (1.5, 2.0) with a width of 5 inches, preserving its aspect ratio."*
4. *"Replace the placeholder rectangle on slide 2 with team_photo.jpg while keeping its exact position and size."*
5. *"Move the entire metric card on slide 2 right by 0.5 inches along with all contained labels and numbers."*
6. *"Make slides 4, 5, and 6 match the header, footer, and category banner of slide 3."*
7. *"Update the stepper on slide 5 so that 'CONFIGURE' is the active step and 'CONNECT' is completed."*
8. *"Validate slides 3, 4, 5, and 6 for overlaps, off-slide objects, or table overflow."*
9. *"Render slides 3, 4, and 5 to verify visual balance and show me a visual diff against the original."*
10. *"Save the changes as updated_platform_deck.pptx."*

---

## 12. Troubleshooting & FAQ

### PowerPoint COM Process Management on Windows
- The server employs process-wide file-based locking (`.ppt_com.lock`), strict STA threading (`pythoncom.CoInitialize()`), 60-second timeouts, and automatic orphan cleanup.
- Individual PowerPoint instances opened by the MCP are cleanly terminated with `Presentation.Close()` and `Application.Quit()` inside `try/finally` blocks, leaving user-opened PowerPoint sessions unaffected.

### Headless Environments & LibreOffice Fallback
- In headless CI/CD environments or Linux servers without Microsoft Office, install LibreOffice:
  ```bash
  sudo apt-get install libreoffice
  ```
- Set `PPT_RENDERER=libreoffice` or `PPT_RENDERER=auto` in your environment.

### Fonts and Typography Drift
- If rendered slides exhibit missing custom fonts, install the TrueType/OpenType font files on the host machine. When proprietary fonts are missing, LibreOffice and PowerPoint will substitute standard system fonts (Calibri, Arial).

---

## License

MIT License.