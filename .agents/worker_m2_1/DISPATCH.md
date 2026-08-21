## 2026-08-21T06:17:19Z
You are the M2 Geometry & Manipulation Worker for the PowerPoint MCP Server project.
Your Working Directory: C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_m2_1

MANDATORY FIRST STEP: Read the following files:
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\ORIGINAL_REQUEST.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\PROJECT.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_core_1\handoff.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

Your exclusive write ownership:
- src/powerpoint_mcp/pptx/geometry.py
- src/powerpoint_mcp/pptx/editor.py
- src/powerpoint_mcp/pptx/ooxml.py
- tests/test_geometry.py
- tests/test_editing.py
- tests/test_text.py
- tests/test_ooxml.py

Implementation Tasks:
1. `src/powerpoint_mcp/pptx/geometry.py`:
   - Align shapes: `align_shapes(shapes, alignment: AlignmentType)` for LEFT, CENTER, RIGHT, TOP, MIDDLE, BOTTOM.
   - Distribute shapes: `distribute_shapes(shapes, mode: DistributionMode, spacing: SpacingMode)` for HORIZONTAL and VERTICAL, supporting EQUAL_GAPS and EQUAL_CENTERS.
   - Equalize dimensions: `equalize_dimensions(shapes, equalize_width=True, equalize_height=True, target_width_inches=None, target_height_inches=None)`.
   - Collision & overlap math: `check_bounding_box_collision(b1, b2, tolerance_emu=0)`, `calculate_overlap_box(b1, b2)`, `calculate_overlap_area(b1, b2)`, `detect_slide_overlaps(slide_model, min_overlap_area_sq_in=0.01)`.
   - Boundary checks: `detect_off_slide_shapes(slide_model, slide_width_inches, slide_height_inches, tolerance_inches=0.01)`.
2. `src/powerpoint_mcp/pptx/editor.py`:
   - `modify_shape(slide_or_prs, slide_number, shape_id, x=None, y=None, width=None, height=None, rotation=None, z_order=None)`:
     * Updates geometry in EMUs without floating drift.
     * Z-order manipulation: reorders `<p:sp>` / `<p:pic>` nodes inside `<p:spTree>`.
   - `modify_text(slide_or_prs, slide_number, shape_id, text=None, font_family=None, font_size=None, bold=None, italic=None, underline=None, color=None, alignment=None, paragraph_spacing=None, line_spacing=None, margins=None)`:
     * Run-level style preservation: If text is replaced, preserves the primary run's typography (font name, size, bold, italic, color).
     * If targeted style parameters are provided (e.g. font_family="Aptos"), applies updates cleanly across all paragraphs/runs or targeted text.
   - `copy_shape(slide_or_prs, source_slide_number, shape_id, target_slide_number=None, offset_x_inches=0.2, offset_y_inches=0.2)`:
     * Deep-copies XML element, assigns new unique shape ID (`<p:cNvPr id="...">`), duplicates relationship references (`r:embed` in `.rels`) for pictures/media.
   - `move_shape(slide_or_prs, slide_number, shape_id, delta_x_inches=None, delta_y_inches=None, x_inches=None, y_inches=None)`:
     * Supports both relative delta and absolute positioning.
   - `resize_shape(slide_or_prs, slide_number, shape_id, width_inches=None, height_inches=None, scale_x=None, scale_y=None)`:
     * Supports absolute dimensions and scaling factors.
   - `delete_shape(slide_or_prs, slide_number, shape_id)`:
     * Removes shape cleanly from `<p:spTree>`.
3. `src/powerpoint_mcp/pptx/ooxml.py`:
   - Purpose-built helpers for OOXML manipulations not directly exposed by python-pptx:
     * `set_shape_transparency(shape, transparency_percent)`
     * `set_gradient_fill(shape, stops)`
     * `set_drop_shadow(shape, ...)`
     * `get_raw_shape_xml(shape)` and `safe_modify_xml(element, modifier_fn)` with validation.
4. Comprehensive Tests:
   - `tests/test_geometry.py` (alignment, distribution, equalization, overlap detection, off-slide checks)
   - `tests/test_editing.py` (modify_shape, move, resize, delete, copy with relations, z-order)
   - `tests/test_text.py` (modify_text, run-level style preservation, font replacement, color changes)
   - `tests/test_ooxml.py` (transparency, gradient helpers, XML safety)
   - Verify with `.venv\Scripts\pytest.exe tests/test_geometry.py tests/test_editing.py tests/test_text.py tests/test_ooxml.py -v`.
5. Document all code and test results in C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_m2_1\handoff.md.
6. Send a brief message back to parent when complete referencing the file path.
