## 2026-08-21T06:00:18Z
You are the E2E Test Infrastructure Worker for the PowerPoint MCP Server project.
Your Working Directory: C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_e2e_1

MANDATORY FIRST STEP: Read the following files:
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\ORIGINAL_REQUEST.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\PROJECT.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\TEST_INFRA.md
- C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\spec_miner_integration_1\handoff.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

Your exclusive write ownership:
- pyproject.toml
- .gitignore
- tests/__init__.py
- tests/conftest.py
- tests/fixtures/__init__.py
- tests/fixtures/create_synthetic_deck.py

Implementation Tasks:
1. `pyproject.toml`:
   - Setup project configuration with name = "powerpoint-mcp", version = "0.1.0", dependencies (python-pptx, mcp, pydantic, pywin32, pillow, numpy, lxml, pytest, pytest-asyncio).
   - Build system standard configuration (setuptools / hatchling / flit_core).
2. `.gitignore`:
   - Ignore .venv, __pycache__, .pytest_cache, .ppt-agent, *.pyc, build, dist, *.egg-info, etc.
3. `tests/fixtures/create_synthetic_deck.py`:
   - Programmatic generator script creating `tests/fixtures/synthetic_sample.pptx` with 3 slides:
     * Slide 1: Title ("Quarterly Performance Overview"), Subtitle ("Q3 2026 Executive Summary"), 3 horizontal KPI boxes with distinct text and colors, 1 image/picture placeholder shape.
     * Slide 2: Title ("Operational Architecture"), 2-column layout (Left Column: Key Initiatives, Right Column: Milestones), 1 diagram shape group (3 interconnected process boxes), footer text.
     * Slide 3: Title ("Audit & Compliance Issues"), intentionally created defects for validation:
       - 2 heavily overlapping shapes (Box A and Box B overlapping by >0.5 inches)
       - 1 shape clipped/extending beyond the right slide boundary
       - 1 shape with suspiciously tiny font (<8pt)
       - 1 text box with large text causing text overflow condition
   - When run as a script (`.venv\Scripts\python.exe tests/fixtures/create_synthetic_deck.py`), it creates `tests/fixtures/synthetic_sample.pptx` if not already present.
4. `tests/conftest.py`:
   - Define pytest fixtures: `synthetic_deck_path`, `sample_presentation`, `temp_workspace_dir`, etc.
5. Run `.venv\Scripts\python.exe tests/fixtures/create_synthetic_deck.py` and verify that the file `tests/fixtures/synthetic_sample.pptx` is generated cleanly.
6. Document implementation and verification in C:\Users\atharva.bedekar\PycharmProjects\Powerpoint MCP\.agents\worker_e2e_1\handoff.md.
7. Send a brief message back to parent when complete referencing the file path.
