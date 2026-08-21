# PowerPoint MCP Server for Antigravity

A production-quality local Model Context Protocol (MCP) server in Python enabling conversational, deterministic PowerPoint (`.pptx`) inspection, surgical editing, high-resolution rendering, visual diffing, and rule-based validation for Google Antigravity.

---

## 1. Overview & Philosophy

Modern AI slide editing often destroys fine-grained formatting, breaks font hierarchies, and resets complex layouts because LLMs attempt to regenerate whole presentations from scratch.

The **PowerPoint MCP Server** implements a deterministic, minimal-diff editing philosophy:
- **Inspect Before Mutating**: Semantic roles (`title`, `subtitle`, `body`, `diagram`, `image`, `footer`) and exact EMU/inch bounding boxes are extracted directly from OpenXML trees.
- **Run-Level Style Preservation**: Text edits preserve font faces, point sizes, colors, margins, and multi-run rich text formatting.
- **Non-Destructive Working Copies**: Modifications occur in isolated session workspaces under `.ppt-agent/sessions/<session_id>/working.pptx` with automatic timestamped backups (`presentation.backup-YYYYMMDD-HHMMSS.pptx`).
- **Visual Verification Loops**: Integrated Windows PowerPoint COM automation and headless LibreOffice rendering generate PNG slide images and pixel difference heatmaps to visually verify layout changes.

---

## 2. Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                      Antigravity CLI Agent                      │
│        (Skill: .agents/skills/powerpoint-editor/SKILL.md)       │
└───────────────────────────────┬─────────────────────────────────┘
                                │ JSON-RPC 2.0 via Stdio
┌───────────────────────────────▼─────────────────────────────────┐
│              PowerPoint MCP Server (server.py)                  │
│       19 FastMCP / MCPServer Tools  &  3 MCP Resources          │
└───────┬───────────────────────┼─────────────────────────┬───────┘
        │                       │                         │
┌───────▼──────────────┐ ┌──────▼───────────────┐  ┌──────▼───────────────┐
│ Inspection & Models  │ │  Editing & Geometry  │  │ Rendering & Diffing  │
│ - Shape Hierarchies  │ │ - Move / Resize      │  │ - Windows PPT COM    │
│ - Semantic Inference │ │ - Align / Distribute │  │ - LibreOffice CLI    │
│ - Multi-Factor Match │ │ - Text Run Styling   │  │ - Pixel Diff Heatmaps│
│ - Rule Validation    │ │ - Safe OOXML Helpers │  │ - Bounding Clusters  │
└──────────────────────┘ └──────────────────────┘  └──────────────────────┘
                                │
┌───────────────────────────────▼─────────────────────────────────┐
│                  Session & Safety Layer                         │
│   .ppt-agent/sessions/<session_id>/working.pptx + backups/      │
└─────────────────────────────────────────────────────────────────┘
```

---

## 3. Prerequisites & Requirements

- **Operating System**: Windows 10/11 (for native PowerPoint COM automation), macOS, or Linux.
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

```bash
python -m venv .venv
# On Windows PowerShell:
.venv\Scripts\Activate.ps1
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

---

## 7. Antigravity Skill Installation

The skill definition is located at `.agents/skills/powerpoint-editor/SKILL.md`. Antigravity automatically discovers skills inside `.agents/skills/`.

The skill encodes:
- The **15 Immutable PowerPoint Editing Rules**.
- Structured decision trees for text edits, geometry adjustments, and reference slide matching.
- Tool call optimization and batching rules.

---

## 8. Supported MCP Tools Reference (19 Core Tools)

| Tool Name | Parameters | Description |
|---|---|---|
| `ppt_open` | `presentation_path: str` | Open presentation and initialize isolated session workspace with working copy. |
| `ppt_inspect_presentation` | `presentation_path: Optional[str]` | Inspect presentation metadata, slide count, dimensions, layouts, and titles. |
| `ppt_inspect_slide` | `slide_number: int, presentation_path: Optional[str]` | Inspect slide shapes, EMU/inch coords, semantic roles, text, and styles. |
| `ppt_inspect_shape` | `slide_number: int, shape_id: int, presentation_path: Optional[str]` | Deep inspection of single shape: runs, fonts, colors, margins, line, fill, XML. |
| `ppt_modify_shape` | `slide_number: int, shape_id: int, x, y, width, height, rotation, z_order, dx, dy, dwidth, dheight, drotation, align, distribute, target_shape_ids, presentation_path` | Modify coordinates, dimensions, rotation, z-order, or multi-shape alignment/distribution. |
| `ppt_modify_text` | `slide_number: int, shape_id: int, text, font_family, font_size, bold, italic, underline, color, alignment, margins, paragraph_spacing, line_spacing, presentation_path` | Modify text and styling with run-level style preservation. |
| `ppt_copy_shape` | `slide_number: int, shape_id: int, target_slide_number, x_offset, y_offset, presentation_path` | Clone shape with formatting and relationships onto same or target slide. |
| `ppt_move_shape` | `slide_number: int, shape_id: int, dx, dy, x, y, presentation_path` | Move shape by absolute coordinates or relative offsets in inches. |
| `ppt_resize_shape` | `slide_number: int, shape_id: int, width, height, scale_width, scale_height, lock_aspect_ratio, presentation_path` | Resize shape with absolute dimensions or scaling multipliers. |
| `ppt_delete_shape` | `slide_number: int, shape_id: int, presentation_path` | Remove shape cleanly from slide shape tree. |
| `ppt_modify_ooxml` | `slide_number: int, shape_id, operation, xpath, attributes, xml_fragment, transparency_percent, gradient_start, gradient_end, shadow_blur_pt, presentation_path` | Safe low-level DrawingML manipulation for transparency, gradients, and shadows. |
| `ppt_validate_slide` | `slide_number: int, rules: Optional[List[str]], presentation_path` | Rule-based validation for overlaps (VAL-01), clipping (VAL-02), tiny fonts (VAL-05), overflow (VAL-04). |
| `ppt_render_slide` | `slide_number: int, output_dir, output_path, renderer, dpi, presentation_path` | Render single slide to high-res PNG via PowerPoint COM or LibreOffice. |
| `ppt_render_presentation` | `output_dir, renderer, dpi, presentation_path` | Render all slides in presentation to PNG images. |
| `ppt_compare_slides` | `slide_a: int, slide_b: int, match_shapes_flag, render_diff, presentation_path` | Compare geometry, layout, typography, and semantic matches between two slides. |
| `ppt_visual_diff` | `before_image: str, after_image: str, output_diff_path, threshold` | Pixel-level difference overlay, changed bounding boxes, and similarity percentage. |
| `ppt_save` | `presentation_path: Optional[str]` | Save working copy back to original path with automatic timestamped backup. |
| `ppt_save_as` | `output_path: str, overwrite: bool, presentation_path` | Save working copy to new destination path. |
| `ppt_revert` | `target: str, presentation_path` | Revert working copy to original presentation or specified backup timestamp. |

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

---

## 11. Example Conversational Commands

Here are typical natural-language prompts handled by Antigravity using the PowerPoint MCP Server:

1. *"Open pitch_deck.pptx and inspect slide 2."*
2. *"Move the title on slide 1 to the left by 0.2 inches."*
3. *"Distribute the three feature boxes on slide 1 horizontally with equal gaps."*
4. *"Align the three cards on slide 2 along their top edges."*
5. *"Change the body text in shape 4 to 'Enterprise Cloud Solutions' and make it bold while keeping the 14pt Calibri font."*
6. *"Duplicate the diagram card on slide 2 and shift it 0.5 inches down."*
7. *"Make slide 2 match the layout and font hierarchy of slide 1, but keep slide 2's content intact."*
8. *"Validate slide 3 for overlaps or off-slide elements and fix any issues found."*
9. *"Render slide 1 and show me a visual diff against the previous version."*
10. *"Save the modified deck as final_presentation.pptx."*

---

## 12. Troubleshooting & FAQ

### PowerPoint COM Errors on Windows
- **Symptom**: `pywintypes.com_error: (-2147417848, 'The object invoked has disconnected from its clients.')`
- **Solution**: Ensure no hung `POWERPNT.EXE` instances exist in Windows Task Manager. Close existing background PowerPoint processes or reboot. The server automatically uses thread-safe COM dispatch with `CoInitialize` / `CoUninitialize`.

### Headless Environments & LibreOffice Fallback
- In headless CI/CD environments or Linux servers without Microsoft Office, install LibreOffice:
  ```bash
  sudo apt-get install libreoffice
  ```
- Set `PPT_RENDERER=libreoffice` or `PPT_RENDERER=auto` in your environment.

### Fonts and Typography Drift
- If rendered slides exhibit missing custom fonts, install the TrueType/OpenType font files on the host machine. If proprietary fonts are missing, LibreOffice and PowerPoint will substitute standard system fonts (Calibri, Arial).

---

## License

MIT License.
