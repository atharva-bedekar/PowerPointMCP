## 2026-08-21T06:17:19Z
You are the M3 Rendering & Verification Worker for the PowerPoint MCP Server project.
Your Working Directory: C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_m3_1

MANDATORY FIRST STEP: Read the following files:
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\ORIGINAL_REQUEST.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\PROJECT.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_core_1\handoff.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

Your exclusive write ownership:
- src/powerpoint_mcp/rendering/__init__.py
- src/powerpoint_mcp/rendering/renderer.py
- src/powerpoint_mcp/rendering/image_diff.py
- src/powerpoint_mcp/rendering/visual_compare.py
- tests/test_rendering.py

Implementation Tasks:
1. `src/powerpoint_mcp/rendering/renderer.py`:
   - `BaseRenderer` abstract base class with `render_slide(presentation_path, slide_number, output_path, width=1920, height=1080)` and `render_presentation(presentation_path, output_dir, width=1920, height=1080)`.
   - `PowerPointRenderer`:
     * Native Windows COM automation using `win32com.client.DispatchEx('PowerPoint.Application')`.
     * Strict STA lifecycle management (`pythoncom.CoInitialize()`, `try...finally`, `presentation.Close()`, `ppt_app.Quit()`, `pythoncom.CoUninitialize()`).
     * Runs invisibly (`WithWindow=False`).
     * Clean error recovery to ensure zero orphaned `POWERPNT.EXE` processes.
   - `LibreOfficeRenderer`:
     * Headless `soffice` export fallback.
     * Checks `shutil.which('soffice')` or standard installation paths.
   - `get_available_renderer(preferred="auto")` -> `BaseRenderer` returning renderer instance with `.renderer_name` and `.is_available`.
2. `src/powerpoint_mcp/rendering/image_diff.py`:
   - `visual_diff(image_a_path, image_b_path, diff_output_path=None, threshold=25)` -> `VisualDiffResult`:
     * Uses numpy pixel matrix subtraction and Pillow compositing.
     * Computes similarity metrics: `similarity_percentage` (0.0 to 100.0%), `pixel_diff_count`, `total_pixels`, `mse`, `psnr`.
     * Detects changed bounding regions using grid block clustering (e.g. 32x32 blocks or connected components) returning list of changed bounding boxes `[{"x": ..., "y": ..., "width": ..., "height": ...}]`.
     * Generates visual diff image with magenta highlight overlay (`#FF00FF`) over modified areas.
3. `src/powerpoint_mcp/rendering/visual_compare.py`:
   - `compare_slides(slide_a_model, slide_b_model, slide_a_img_path=None, slide_b_img_path=None)` -> `SlideComparisonResult`:
     * Compares geometric layout, shape counts, typography, and positions.
     * Incorporates visual diff metrics if rendered images are provided.
4. Comprehensive Tests (`tests/test_rendering.py`):
   - Test renderer detection (`get_available_renderer`).
   - Test PowerPoint COM rendering on Windows with `synthetic_sample.pptx` (generates real PNGs, verifies dimensions and valid image header).
   - Test LibreOffice fallback detection & error handling.
   - Test `visual_diff` on identical images (100% similarity, 0 diffs).
   - Test `visual_diff` on edited/modified images (detects bounding box of changes, computes similarity < 100%, generates magenta diff image).
   - Test `compare_slides` geometric and visual comparison.
   - Verify with `.venv\Scripts\pytest.exe tests/test_rendering.py -v`.
5. Document all code and test results in C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_m3_1\handoff.md.
6. Send a brief message back to parent when complete referencing the file path.
