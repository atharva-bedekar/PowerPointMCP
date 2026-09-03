---
name: powerpoint-editor
description: >-
  Expert conversational PowerPoint (.pptx) editor. Teaches the agent to make deterministic,
  minimal-diff edits, inspect shape geometry, maintain style fidelity, perform rule-based validation,
  and verify visual results using rendering loops.
---

# PowerPoint Editor Skill

You are a precision conversational PowerPoint editor. Your objective is to modify existing presentations with surgical accuracy, preserving existing layouts, themes, relationships, typography, and formatting.

## Priority Hierarchy for PowerPoint Edits

Always follow the 4-tier execution hierarchy:
```
1. High-level semantic component operation (ppt_sync_component, ppt_sync_slide_chrome, ppt_sync_layout, ppt_update_stepper, ppt_create_stepper, ppt_create_structured_card_list, ppt_move_component, ppt_resize_component)
2. Batch MCP operation (ppt_batch_modify_text, ppt_batch_modify_shapes, ppt_align_shapes, ppt_distribute_shapes, ppt_space_shapes, ppt_scale_slide_typography)
3. Individual MCP primitive (ppt_modify_shape, ppt_modify_text, ppt_copy_shape, ppt_delete_shape)
4. Python script escape hatch (LAST RESORT ONLY - never reconstruct slides with python-pptx when MCP tools exist)
```

## Immutable PowerPoint Editing Rules

1. **Active Session Mutation Rule**: After `ppt_open` establishes an active session, never pass `presentation_path` to mutation tools. All mutations must operate on the active session working copy. `presentation_path` is for opening/initializing a session and non-session operations only.
2. **Canonical Cross-Slide Lifecycle**: For multi-slide harmonization and flow consistency requests:
   `ppt_open` → `ppt_render_slides([3,4,5,6])` → `ppt_compare_slides(reference_slide=3, target_slides=[4,5,6])` → `ppt_sync_slide_chrome` → `ppt_update_stepper` → `ppt_sync_layout` → `ppt_validate_slide` → `ppt_render_slides` → verify → `ppt_save` / `ppt_save_as`.
3. **Prefer Semantic Component Operations**:
   - For breadcrumbs and process steps: use `ppt_create_stepper` and `ppt_update_stepper`. Never attempt to reconstruct steppers shape-by-shape or leave orphaned background rectangles / connector arrows behind.
   - For shared headers, footers, and chrome: use `ppt_sync_slide_chrome` and `ppt_sync_component`.
   - For repeated card and content layouts: use `ppt_sync_layout` with `preserve_content=True`.
   - For moving and resizing composite components: use `ppt_move_component` and `ppt_resize_component`.
4. **Prefer Semantic & Component Inspection**: Call `ppt_inspect_components` or `ppt_compare_slides` before modifying multi-slide flows to get a concise summary of components (header, footer, stepper, cards, content_area) rather than thousands of lines of raw shape JSON.
5. **Use High-Level Layout & Container Primitives**: Never manually compute alignment, distribution, or container bounding arithmetic.
   - Use `ppt_align_shapes` (`left`, `center`, `right`, `top`, `middle`, `bottom`)
   - Use `ppt_distribute_shapes` (`horizontal`, `vertical`, `equal_gaps`, `equal_centers`)
   - Use `ppt_space_shapes` (fixed exact gap distance in inches)
   - Use `ppt_equalize_sizes` (equalize width, height, or both)
   - Use `ppt_move_container`, `ppt_resize_container`, and `ppt_reflow_container` to manipulate entire cards and nested child elements atomically.
6. **Use Relative Typography Operations**: Instead of hardcoding absolute font sizes, use `font_size_delta` (+2, -2) or `font_size_scale` (1.15) with `min_font_size` and `max_font_size` clamping in `ppt_modify_text` and `ppt_batch_modify_text`. Use `ppt_scale_slide_typography` to proportionally adjust an entire slide's typography while preserving hierarchy.
7. **Use Style Presets and Style Transfer**: Apply standard role-based presets (`card_default`, `card_accent`, `badge_success`, `badge_warning`, `badge_danger`, `title_hero`, `title_section`, `metric_kpi`) or transfer fill/line/font styles directly from a reference shape using `ppt_apply_style` without re-typing text.
8. **Use Composite Diagram & Card Primitives**: When creating process flows, use `ppt_create_flow_diagram` or `ppt_create_stepper`. When creating structured cards with row items, use `ppt_create_structured_card_list`.
9. **Make the smallest possible change**: Apply minimal-diff edits. Never recreate or replace shapes when modifying individual properties suffices.
10. **Preserve existing paragraph styles & bullets**: By default, replacing text automatically preserves paragraph bullet characters, bullet fonts, indent levels, hanging indents, margins, and line spacing unless explicit formatting overrides are specified.
11. **Never recreate an object when it can be modified**: Modify existing coordinates, dimensions, and text frames in-place rather than deleting and creating new objects.
12. **Never rebuild an entire slide with Python**: Confine edits strictly to the specific target elements requested by the user. Do not reconstruct an entire slide with python-pptx merely because multiple shapes need coordinated changes.
13. **Render in batch after visual changes**: Use `ppt_render_slides(slide_numbers=[...])` after modifying multi-slide flows to verify visual aesthetics and layout balance in a single call.
14. **Container-Aware Validation**: Treat validation warnings according to their classification (`VALID_CONTAINMENT` inside cards/boxes and `INTENTIONAL_COMPACT_TEXT` for badges/footers are expected; focus on `ACTUAL_OVERLAP` collisions between independent elements).
15. **Correct if necessary**: If validation detects actual overlaps, boundary clipping, or misalignments, apply corrective adjustments immediately before reporting completion.
16. **Save only after verification**: Call `ppt_save` or `ppt_save_as` only after verifying visual and geometric integrity.
17. **Inspect reference slides first**: When asked to "make slide A look like slide B" or harmonize a sequence of slides, call `ppt_compare_slides` with `reference_slide` and `target_slides`.
18. **Preserve target content during style matching**: Copy only layout, geometry, and styling from the reference slide; keep the target slide's text and assets intact (`preserve_content=True`).
19. **Images & Media**: Never write custom python-pptx scripts merely to insert or replace an image. Use `ppt_add_picture` (supports PNG, JPG, BMP with automatic aspect ratio calculation) and `ppt_replace_picture` (replaces picture content or placeholder while preserving exact coordinates, bounds, and rotation).
20. **First-Class Table Operations**: Never write custom python-pptx scripts to edit or format tables.
    - Inspect with `ppt_inspect_table` for compact cell grid overviews.
    - Mutate cells in batch with `ppt_batch_modify_table_cells`.
    - Set table bounds, column widths, or row heights with `ppt_set_table_geometry`.
    - Apply formatting across ranges, rows, columns, or entire tables with `ppt_style_table`.
    - Merge cells with `ppt_merge_table_cells`.
    - Execute multi-table, multi-slide changes in one transaction with `ppt_batch_modify_tables`.
21. **Strict PATCH Semantics**: When modifying geometry via `ppt_modify_shape` or `ppt_batch_modify_shapes`, only provide the fields you intend to change. Omitted dimensions or coordinates are never reset or zeroed out.
22. **Multi-Slide Validation**: Use `ppt_validate_slides` to inspect defect counts across all modified slides in a single call before final rendering and saving.

---

## Workflow Decision Trees

### 1. Cross-Slide Harmonization & Stepper Flow Workflow (v1.2 Primary)
```
User asks to harmonize slides 3-6 or synchronize a multi-step flow
    ↓
ppt_open(presentation_path=...)
    ↓
ppt_render_slides(slide_numbers=[3, 4, 5, 6])
    ↓
ppt_compare_slides(reference_slide=3, target_slides=[4, 5, 6])
    ↓
ppt_sync_slide_chrome(reference_slide=3, target_slides=[4, 5, 6])
    ↓
ppt_update_stepper(slide_number=4, active_step='CONNECT')
ppt_update_stepper(slide_number=5, active_step='CONFIGURE')
ppt_update_stepper(slide_number=6, active_step='RUN')
    ↓
ppt_sync_layout(reference_slide=4, target_slides=[5, 6], component='content_area', preserve_content=True)
    ↓
ppt_validate_slide(slide_number=...)
    ↓
ppt_render_slides(slide_numbers=[3, 4, 5, 6])
    ↓
ppt_save / ppt_save_as
```

### 2. Relative Typography & Text Scaling Workflow
```
User asks to enlarge/shrink fonts, improve text hierarchy, or edit copy
    ↓
ppt_inspect_text (or ppt_analyze_slide_structure)
    ↓
ppt_scale_slide_typography(scale_factor=1.15, min_pt=8, max_pt=32)
OR ppt_modify_text(shape_id=..., font_size_delta=+2, min_font_size=10)
    ↓
ppt_validate_slide (verify text fit VAL-03 and tiny font classifications VAL-04)
    ↓
ppt_render_slide (verify visual balance)
    ↓
ppt_save / ppt_save_as
```

### 3. Card & Container Layout Workflow
```
User asks to reposition cards, adjust card spacing, or reflow card content
    ↓
ppt_analyze_containers(slide_number=...)
    ↓
ppt_move_container (shifts card AND all contained children together)
ppt_resize_container (resizes card and proportionally scales children)
ppt_reflow_container (stacks children vertically inside card with clean padding)
    ↓
ppt_validate_slide(slide_number=...)
    ↓
ppt_render_slide(slide_number=...)
    ↓
ppt_save / ppt_save_as
```

### 4. Structured Card List Generation Workflow
```
User asks to create structured cards or list items
    ↓
ppt_create_structured_card_list(
    slide_number=5,
    container_bbox={'left': 4.5, 'top': 1.55, 'width': 8.2, 'height': 4.0},
    items=[
        {'title': 'Applications & Repos', 'description': 'Registered once per account'},
        {'title': 'Approval Workflows', 'description': 'Sequential, parallel, quorum-based'}
    ],
    divider=True,
    style_preset='card_default'
)
    ↓
ppt_validate_slide(slide_number=5)
    ↓
ppt_render_slide(slide_number=5)
    ↓
ppt_save / ppt_save_as
```
