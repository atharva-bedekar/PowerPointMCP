## 2026-08-21T05:53:00Z
You are the Integration & Safety Spec Miner for the PowerPoint MCP Server project.
Your Working Directory: C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_integration_1

MANDATORY FIRST STEP: Read C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\ORIGINAL_REQUEST.md and C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\Build a PowerPoint Editing MCP Server for Antigravity.md.

Your mission:
1. Extract exhaustive, precise technical specifications for R3 (Session & Safety Layer) and R4 (Antigravity Integration & Tooling):
   - Session & Working Copy Model: .ppt-agent/ directory structure, session isolation, working.pptx lifecycle, timestamped backups (presentation.backup-YYYYMMDD-HHMMSS.pptx), ppt_open, ppt_save, ppt_save_as, ppt_revert.
   - Rule-based Slide Validation (ppt_validate_slide): overlap detection, off-slide / boundary clipping, text overflow heuristics, tiny fonts, inconsistent title positions/margins, duplicate objects, extreme rotations. Structured warnings and diagnostic output.
   - MCP Server & Tools: FastMCP / MCP Python SDK stdio server setup, full list of 19 tools (ppt_open, ppt_inspect_presentation, ppt_inspect_slide, ppt_inspect_shape, ppt_modify_shape, ppt_modify_text, ppt_copy_shape, ppt_move_shape, ppt_resize_shape, ppt_delete_shape, ppt_modify_ooxml, ppt_validate_slide, ppt_render_slide, ppt_render_presentation, ppt_compare_slides, ppt_visual_diff, ppt_save, ppt_save_as, ppt_revert), tool schemas, descriptions, structured return formats, structured error handling (ShapeNotFound, SlideNotFound, etc.).
   - MCP Resources: ppt://current/presentation, ppt://current/slide/{slide_number}, ppt://current/slide/{slide_number}/render.
   - CLI Debugging Utilities: scripts/inspect_pptx.py and scripts/render_pptx.py (arguments, stdout format, exit codes).
   - Antigravity Configuration: .agents/mcp_config.json (merging safely with any existing config, stdio transport, python/uv command).
   - Antigravity Skill: .agents/skills/powerpoint-editor/SKILL.md (YAML frontmatter, 15 editing rules, decision tree, inspect-reason-modify-render-verify loop).
   - Synthetic 3-Slide Test Presentation: exact specification for Slide 1 (title, subtitle, 3 boxes, image placeholder), Slide 2 (title, 2-column, diagram, footer), Slide 3 (title, intentional overlaps).
   - Test Suite Specifications: pytest test cases for all modules and in-memory MCP client testing.
   - README.md: full outline and required sections.
2. Structure your report with exact JSON schemas, configurations, CLI syntax, and safety protocols.
3. Write your complete spec handoff to C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_integration_1\handoff.md.
4. Send a brief message back to parent when complete referencing the file path.
