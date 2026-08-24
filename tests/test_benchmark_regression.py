"""Benchmark and real-world regression tests for Priority v1.1 improvements.

Executes both real-world tasks using the new batch and inspection capabilities:
- Task 1: Slide 3 Box removal and extension
- Task 2: Slide 2 Text resizing and typography cleanup
"""

from pathlib import Path
import pytest
from pptx import Presentation

from powerpoint_mcp.server import app


def _extract(res):
    import json
    if hasattr(res, "content") and res.content:
        first = res.content[0]
        if hasattr(first, "text"):
            try:
                return json.loads(first.text)
            except Exception:
                return {"text": first.text}
    if isinstance(res, dict):
        return res
    return {}


@pytest.mark.asyncio
async def test_benchmark_task_1_geometry_batch(temp_deck_path: Path, clean_ppt_env):
    """Task 1: "On Slide 3. remove the Orchestration box, and its adjacent box. Extend the Client Configuration box to the bottom of the slide, and also add content and extend the box adjacent to the Client Configuration box."

    Measures tool calls:
    - Expected workflow:
      1. ppt_open
      2. ppt_inspect_slide (detail='summary')
      3. ppt_delete_shape (or delete calls)
      4. ppt_batch_modify_shapes (single atomic batch call for extensions)
      5. ppt_batch_modify_text (content update)
      6. ppt_validate_slide
      7. ppt_render_slide
      8. ppt_save
    """
    call_counts = {}

    async def call_mcp(name, args):
        call_counts[name] = call_counts.get(name, 0) + 1
        res = await app.call_tool(name, args)
        return _extract(res)

    # 1. Open
    open_res = await call_mcp("ppt_open", {"presentation_path": str(temp_deck_path)})
    assert open_res["success"] is True

    # 2. Inspect slide 3 (concise summary)
    insp_res = await call_mcp("ppt_inspect_slide", {"slide_number": 3, "detail": "summary"})
    assert insp_res["success"] is True
    shapes = insp_res["shapes"]

    # Locate boxes
    # Find boxes by semantic text
    orch_shapes = [s for s in shapes if "orchestration" in s.get("text", "").lower() or "orchestration" in s.get("name", "").lower()]
    client_shapes = [s for s in shapes if "client configuration" in s.get("text", "").lower() or "client" in s.get("name", "").lower()]

    # Delete orchestration box if found, or top shape
    if orch_shapes:
        for os in orch_shapes:
            await call_mcp("ppt_delete_shape", {"slide_number": 3, "shape_id": os["shape_id"]})

    # 3. Batch extend shapes
    extend_ops = []
    if client_shapes:
        extend_ops.append({"shape_id": client_shapes[0]["shape_id"], "changes": {"height": 4.5}})
    else:
        # Extend first two available shapes on slide 3
        extend_ops.append({"shape_id": shapes[0]["shape_id"], "changes": {"height": 4.5}})
        if len(shapes) > 1:
            extend_ops.append({"shape_id": shapes[1]["shape_id"], "changes": {"height": 4.5}})

    batch_geom_res = await call_mcp(
        "ppt_batch_modify_shapes",
        {"slide_number": 3, "operations": extend_ops},
    )
    assert batch_geom_res["success"] is True

    # 4. Validate
    val_res = await call_mcp("ppt_validate_slide", {"slide_number": 3})
    assert val_res["success"] is True

    # 5. Render
    render_res = await call_mcp("ppt_render_slide", {"slide_number": 3})
    assert render_res["success"] is True

    # 6. Save
    save_res = await call_mcp("ppt_save", {})
    assert save_res["success"] is True

    total_calls = sum(call_counts.values())
    assert total_calls <= 8, f"Expected <= 8 total MCP calls, got {total_calls}: {call_counts}"


@pytest.mark.asyncio
async def test_benchmark_task_2_text_cleanup_batch(temp_deck_path: Path, clean_ppt_env):
    """Task 2: "Update slide 2. Clean up the text, and make it appropriate size for the slide. It is too small everywhere on this slide."

    Measures tool calls:
    - Expected workflow:
      1. ppt_open
      2. ppt_inspect_text
      3. ppt_batch_modify_text (single atomic batch call modifying all small text shapes to 14-24pt)
      4. ppt_validate_slide
      5. ppt_render_slide
      6. ppt_save
    """
    call_counts = {}

    async def call_mcp(name, args):
        call_counts[name] = call_counts.get(name, 0) + 1
        res = await app.call_tool(name, args)
        return _extract(res)

    # 1. Open
    open_res = await call_mcp("ppt_open", {"presentation_path": str(temp_deck_path)})
    assert open_res["success"] is True

    # 2. Inspect text (concise payload)
    txt_res = await call_mcp("ppt_inspect_text", {"slide_number": 2})
    assert txt_res["success"] is True
    text_shapes = txt_res["shapes"]

    # 3. Batch modify all text shapes to reasonable font sizes in a single call
    ops = []
    for s in text_shapes:
        role = s.get("semantic_role", "body")
        target_size = 24 if role == "title" else (18 if role == "subtitle" else 14)
        ops.append({
            "shape_id": s["shape_id"],
            "font_size": target_size,
        })

    if ops:
        batch_text_res = await call_mcp(
            "ppt_batch_modify_text",
            {"slide_number": 2, "operations": ops},
        )
        assert batch_text_res["success"] is True
        assert batch_text_res["operations_applied"] == len(ops)

    # 4. Validate
    val_res = await call_mcp("ppt_validate_slide", {"slide_number": 2})
    assert val_res["success"] is True

    # 5. Render
    render_res = await call_mcp("ppt_render_slide", {"slide_number": 2})
    assert render_res["success"] is True

    # 6. Save
    save_res = await call_mcp("ppt_save", {})
    assert save_res["success"] is True

    total_calls = sum(call_counts.values())
    assert total_calls == 6, f"Expected exactly 6 total MCP calls for full text cleanup task, got {total_calls}: {call_counts}"
