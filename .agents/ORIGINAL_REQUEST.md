# Original User Request

## Initial Request — 2026-08-21T05:51:15Z

Build a production-quality local Model Context Protocol (MCP) server in Python that enables conversational, deterministic PowerPoint (.pptx) inspection, precise editing, visual rendering, and validation for the Antigravity CLI agent.

Working directory: C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP
Integrity mode: development

## Requirements

### R1. Deterministic PPTX Inspection & Geometry Engine
Expose structured MCP tools to inspect presentation metadata, slide elements, and shapes with precise EMU/inch coordinate mapping and conservative semantic role detection (e.g. title, body, diagram). Implement deterministic shape manipulation (move, resize, align, distribute, copy, delete), text editing (with run-level style preservation), and safe OOXML fallback helpers.

### R2. Rendering & Visual Verification Pipeline
Implement rendering capabilities supporting Windows PowerPoint COM automation with graceful LibreOffice headless fallback. Implement deterministic visual comparison and diffing (pixel diff, changed bounding regions, similarity metrics) between slides or before/after edits.

### R3. Session & Safety Layer
Implement a non-destructive session/working-copy model with automatic timestamped backups before modifications (`ppt_save`, `ppt_save_as`, `ppt_revert`) to ensure source files are never overwritten unintentionally. Include rule-based slide validation (`ppt_validate_slide`) for overlaps, boundary clipping, and text overflow.

### R4. Antigravity Integration & Tooling
Expose standard MCP stdio server tools and resources (`ppt://current/...`). Provide standalone CLI debugging scripts (`inspect_pptx.py`, `render_pptx.py`), Antigravity workspace configuration (`.agents/mcp_config.json`), and an Antigravity skill (`.agents/skills/powerpoint-editor/SKILL.md`) guiding agent behavior through inspect-modify-render-verify loops.

## Acceptance Criteria

### Automated Test Suite & Tool Verification
- [ ] Test suite passing with pytest covering inspection, geometry, text modification, copy/delete, OOXML helpers, rendering detection, and MCP tool invocations via an in-memory MCP client.
- [ ] Programmatic synthetic 3-slide test presentation generated and used for integration testing.

### End-to-End Workflow Verification
- [ ] End-to-end multi-step verification passing:
  1. Slide inspection returning structured dimensions and semantic roles.
  2. Exact coordinate modification (e.g., move title 0.2 inches left).
  3. Shape alignment and horizontal distribution.
  4. Cross-slide visual/geometric matching preserving content.
  5. Slide validation detecting deliberate overlaps.
  6. Slide rendering generating valid PNG image output.
  7. Non-destructive save-as to output path.

### Antigravity Discovery & Packaging
- [ ] `.agents/mcp_config.json` correctly configured for local stdio execution.
- [ ] `.agents/skills/powerpoint-editor/SKILL.md` present and adheres to Antigravity YAML frontmatter format.
- [ ] Complete `README.md` with architecture, setup, tool reference, and troubleshooting instructions.
