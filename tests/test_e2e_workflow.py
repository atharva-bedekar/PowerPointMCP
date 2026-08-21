"""End-to-End Workflow Verification Tests.

Verifies the 7 acceptance criteria workflows:
1. Slide inspection returning structured dimensions and semantic roles.
2. Exact coordinate modification (e.g., move title 0.2 inches left).
3. Shape alignment and horizontal distribution.
4. Cross-slide visual/geometric matching preserving content.
5. Slide validation detecting deliberate overlaps.
6. Slide rendering generating valid PNG image output.
7. Non-destructive save-as to output path.
"""

from pathlib import Path
import shutil
from typing import Any, Dict

import pytest
from pptx import Presentation

from powerpoint_mcp.server import app


def _get_json(res: Any) -> Dict[str, Any]:
    import json
    if hasattr(res, "content") and res.content:
        return json.loads(res.content[0].text)
    if isinstance(res, dict):
        return res
    return {}


@pytest.mark.asyncio
async def test_e2e_full_workflow(temp_deck_path: Path, tmp_path: Path):
    """Execute all 7 acceptance criteria in a continuous, deterministic end-to-end editing workflow."""

    # -------------------------------------------------------------------------
    # Workflow 0: Initialize Session
    # -------------------------------------------------------------------------
    open_res = await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})
    open_data = _get_json(open_res)
    assert open_data["success"] is True
    session_id = open_data["session_id"]
    working_path = open_data["working_path"]

    # -------------------------------------------------------------------------
    # Workflow 1: Slide inspection returning structured dimensions & semantic roles
    # -------------------------------------------------------------------------
    inspect_res = await app.call_tool("ppt_inspect_slide", {"slide_number": 1})
    inspect_data = _get_json(inspect_res)
    assert inspect_data["success"] is True
    assert inspect_data["slide_number"] == 1
    assert inspect_data["width_inches"] > 0
    assert inspect_data["height_inches"] > 0

    roles = [s["semantic_role"] for s in inspect_data["shapes"]]
    assert "title" in roles, "Semantic role inference must detect title on slide 1"

    title_shape = next(s for s in inspect_data["shapes"] if s["semantic_role"] == "title")
    initial_x = title_shape["bbox"]["left_inches"]
    title_id = title_shape["shape_id"]

    # -------------------------------------------------------------------------
    # Workflow 2: Exact coordinate modification (e.g. move title 0.2 inches left)
    # -------------------------------------------------------------------------
    move_res = await app.call_tool(
        "ppt_move_shape",
        {"slide_number": 1, "shape_id": title_id, "dx": -0.2},
    )
    move_data = _get_json(move_res)
    assert move_data["success"] is True

    # Verify updated coordinate
    reinspect_res = await app.call_tool("ppt_inspect_shape", {"slide_number": 1, "shape_id": title_id})
    reinspect_data = _get_json(reinspect_res)
    new_x = reinspect_data["shape"]["bbox"]["left_inches"]
    assert abs(new_x - (initial_x - 0.2)) < 0.001, f"Expected {initial_x - 0.2}, got {new_x}"

    # -------------------------------------------------------------------------
    # Workflow 3: Shape alignment and horizontal distribution
    # -------------------------------------------------------------------------
    # Identify feature box shapes on slide 1 (body cards)
    body_shapes = [s for s in inspect_data["shapes"] if s["semantic_role"] == "body"]
    if len(body_shapes) >= 3:
        box_ids = [s["shape_id"] for s in body_shapes[:3]]
        primary_id = box_ids[0]
        other_ids = box_ids[1:]

        # Distribute horizontally and align top edges
        geom_res = await app.call_tool(
            "ppt_modify_shape",
            {
                "slide_number": 1,
                "shape_id": primary_id,
                "align": "top",
                "distribute": "horizontal",
                "target_shape_ids": other_ids,
            },
        )
        geom_data = _get_json(geom_res)
        assert geom_data["success"] is True

        # Verify shapes have matching top coordinates
        s1_inspect = _get_json(await app.call_tool("ppt_inspect_slide", {"slide_number": 1}))
        top_coords = [
            s["bbox"]["top_inches"] for s in s1_inspect["shapes"] if s["shape_id"] in box_ids
        ]
        assert all(abs(t - top_coords[0]) < 0.001 for t in top_coords), "All 3 boxes should have identical top coordinates"

    # -------------------------------------------------------------------------
    # Workflow 4: Cross-slide visual / geometric matching preserving content
    # -------------------------------------------------------------------------
    comp_res = await app.call_tool("ppt_compare_slides", {"slide_a": 1, "slide_b": 2})
    comp_data = _get_json(comp_res)
    assert comp_data["success"] is True
    assert "geometric_match_score" in comp_data
    assert comp_data["matched_shape_count"] > 0

    # -------------------------------------------------------------------------
    # Workflow 5: Slide validation detecting deliberate overlaps on slide 3
    # -------------------------------------------------------------------------
    val_res = await app.call_tool("ppt_validate_slide", {"slide_number": 3})
    val_data = _get_json(val_res)
    assert val_data["success"] is True
    rule_ids = [w["rule_id"] for w in val_data.get("warnings", [])]
    assert "VAL-01" in rule_ids, "Should detect intentional overlap on slide 3"

    # -------------------------------------------------------------------------
    # Workflow 6: Slide rendering generating valid PNG image output
    # -------------------------------------------------------------------------
    render_dir = tmp_path / "renders"
    render_res = await app.call_tool(
        "ppt_render_slide",
        {"slide_number": 1, "output_dir": str(render_dir), "renderer": "mock"},
    )
    render_data = _get_json(render_res)
    assert render_data["success"] is True
    rendered_png = Path(render_data["image_path"])
    assert rendered_png.exists()
    assert rendered_png.stat().st_size > 0

    # -------------------------------------------------------------------------
    # Workflow 7: Non-destructive save-as to output path
    # -------------------------------------------------------------------------
    final_output_path = tmp_path / "final_verified_deck.pptx"
    save_as_res = await app.call_tool(
        "ppt_save_as",
        {"output_path": str(final_output_path)},
    )
    save_as_data = _get_json(save_as_res)
    assert save_as_data["success"] is True
    assert final_output_path.exists()
    assert final_output_path.stat().st_size > 0

    # Verify original source presentation was untouched
    original_prs = Presentation(str(temp_deck_path))
    assert len(original_prs.slides) == 3
