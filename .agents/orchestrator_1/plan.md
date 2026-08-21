# Orchestration Plan: PowerPoint MCP Server

## Objective
Build a production-quality local Model Context Protocol (MCP) server in Python for PowerPoint (.pptx) deterministic inspection, editing, rendering, visual diffing, and validation for Antigravity.

## Orchestration Strategy: Dual Track Project Pattern

### Phase 0: Survey & Environment Discovery
- Spawn 3 Explorers / Spec Miners:
  1. `explorer_env`: Discover local Python environment (Python version, uv, pytest, python-pptx, mcp SDK, pywin32, PowerPoint COM, LibreOffice, Node/npx, etc.).
  2. `spec_miner_core`: Extract all specification requirements for R1 & R2 (Data models, inspection, geometry, manipulation, OOXML, COM/LibreOffice rendering, visual comparison).
  3. `spec_miner_integration`: Extract all specification requirements for R3 & R4 (Session/safety, backups, rule-based validation, MCP tools/resources, CLI utilities, Antigravity skill & mcp_config.json).
- Synthesize into `PROJECT.md` (with complete Feature Inventory & Milestones) and `TEST_INFRA.md`.

### Phase 1: Dual Track Execution
- **E2E Testing Track**:
  - Design test runner & fixtures (including synthetic 3-slide presentation).
  - Implement Tier 1 (Feature coverage ≥5/feature), Tier 2 (Boundaries), Tier 3 (Cross-feature interactions), Tier 4 (Real-world scenarios).
  - Publish `TEST_READY.md`.
- **Implementation Track**:
  - Milestone M1: Core Models & PPTX Inspection Engine (`powerpoint_mcp/models`, `pptx/inspector.py`).
  - Milestone M2: Geometry, Manipulation & OOXML Engine (`pptx/geometry.py`, `pptx/editor.py`, `pptx/ooxml.py`).
  - Milestone M3: Rendering & Visual Verification Pipeline (`rendering/renderer.py`, `image_diff.py`, `visual_compare.py`).
  - Milestone M4: Session, Safety & Validation Engine (`utils/validation.py`, `tools/versioning.py`, working-copy session model).
  - Milestone M5: MCP Server Tools, Resources, CLI & Antigravity Packaging (`server.py`, `tools/*`, `scripts/*`, `SKILL.md`, `mcp_config.json`, `README.md`).

### Phase 2: Final Milestone & Integration
- Phase 2A: 100% Pass on E2E Test Suite (Tiers 1-4).
- Phase 2B: Adversarial Coverage Hardening (Tier 5 with Challenger & test coverage audit).
- Phase 2C: Full Forensic Integrity Audit across all modules.
- Phase 2D: Report results to user and parent.
