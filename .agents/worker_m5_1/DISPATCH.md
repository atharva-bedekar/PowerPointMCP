## 2026-08-21T06:24:55Z
You are the M5 MCP Server & Packaging Worker for the PowerPoint MCP Server project.
Your Working Directory: C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_m5_1

MANDATORY FIRST STEP: Read the following files:
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\ORIGINAL_REQUEST.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\PROJECT.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\TEST_INFRA.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_integration_1\handoff.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

Your exclusive write ownership:
- src/powerpoint_mcp/server.py
- src/powerpoint_mcp/tools/__init__.py
- src/powerpoint_mcp/tools/inspection.py
- src/powerpoint_mcp/tools/editing.py
- src/powerpoint_mcp/tools/rendering.py
- scripts/inspect_pptx.py
- scripts/render_pptx.py
- .agents/mcp_config.json
- .agents/skills/powerpoint-editor/SKILL.md
- README.md
- tests/test_mcp.py
- tests/test_e2e_workflow.py

Implementation Tasks:
1. `src/powerpoint_mcp/tools/`:
   - `inspection.py`: handlers for `ppt_inspect_presentation`, `ppt_inspect_slide`, `ppt_inspect_shape`, `ppt_compare_slides`, `ppt_validate_slide`.
   - `editing.py`: handlers for `ppt_modify_shape`, `ppt_modify_text`, `ppt_copy_shape`, `ppt_move_shape`, `ppt_resize_shape`, `ppt_delete_shape`, `ppt_modify_ooxml`.
   - `rendering.py`: handlers for `ppt_render_slide`, `ppt_render_presentation`, `ppt_visual_diff`.
   - Integrate with `SessionManager` in `versioning.py` for `ppt_open`, `ppt_save`, `ppt_save_as`, `ppt_revert`.
2. `src/powerpoint_mcp/server.py`:
   - `MCPServer("powerpoint-mcp")` registering all 19 MCP tools with rich schemas and docstrings:
     1. `ppt_open`
     2. `ppt_inspect_presentation`
     3. `ppt_inspect_slide`
     4. `ppt_inspect_shape`
     5. `ppt_modify_shape`
     6. `ppt_modify_text`
     7. `ppt_copy_shape`
     8. `ppt_move_shape`
     9. `ppt_resize_shape`
     10. `ppt_delete_shape`
     11. `ppt_modify_ooxml`
     12. `ppt_validate_slide`
     13. `ppt_render_slide`
     14. `ppt_render_presentation`
     15. `ppt_compare_slides`
     16. `ppt_visual_diff`
     17. `ppt_save`
     18. `ppt_save_as`
     19. `ppt_revert`
   - Register MCP Resources:
     * `ppt://current/presentation`
     * `ppt://current/slide/{slide_number}`
     * `ppt://current/slide/{slide_number}/render`
   - Stdio transport runner via `app.run_stdio_async()`.
   - Structured error handling: catch domain exceptions and return `{ "success": False, "error_type": "...", "message": "...", "details": ... }`.
3. Standalone CLI scripts:
   - `scripts/inspect_pptx.py`: CLI wrapper inspecting presentations, outputting formatted summaries or JSON.
   - `scripts/render_pptx.py`: CLI wrapper exporting slides to PNG via COM/LibreOffice.
4. Antigravity Packaging:
   - `.agents/mcp_config.json`: configured for stdio transport with command `uv run python -m powerpoint_mcp.server`.
   - `.agents/skills/powerpoint-editor/SKILL.md`: YAML frontmatter, 15 editing rules, decision tree, inspect-modify-render-verify workflow.
   - `README.md`: complete 12-section production documentation.
5. Tests:
   - `tests/test_mcp.py`: In-memory MCP client testing calling all 19 tools and resources via `app.call_tool` and `app.read_resource`.
   - `tests/test_e2e_workflow.py`: End-to-end integration test verifying all 7 acceptance criteria workflows.
   - Run entire test suite: `.venv\Scripts\pytest.exe -v`.
6. Document all code and test results in C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_m5_1\handoff.md.
7. Send a brief message back to parent when complete referencing the file path.
