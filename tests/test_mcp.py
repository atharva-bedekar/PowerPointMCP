"""In-memory MCP client tests for PowerPoint MCP Server tools and resources."""

import json
from pathlib import Path
from typing import Any, Dict

import pytest

from powerpoint_mcp.server import app


def _extract_tool_result(res: Any) -> Dict[str, Any]:
    """Helper to extract dictionary from CallToolResult."""
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
async def test_mcp_tool_discovery():
    """Verify that all 19 tools are registered on the MCPServer instance."""
    tools = await app.list_tools()
    tool_names = {t.name for t in tools}

    expected_tools = {
        "ppt_open",
        "ppt_inspect_presentation",
        "ppt_inspect_slide",
        "ppt_inspect_shape",
        "ppt_modify_shape",
        "ppt_modify_text",
        "ppt_copy_shape",
        "ppt_move_shape",
        "ppt_resize_shape",
        "ppt_delete_shape",
        "ppt_modify_ooxml",
        "ppt_validate_slide",
        "ppt_render_slide",
        "ppt_render_presentation",
        "ppt_compare_slides",
        "ppt_visual_diff",
        "ppt_save",
        "ppt_save_as",
        "ppt_revert",
    }

    assert expected_tools.issubset(tool_names), f"Missing tools: {expected_tools - tool_names}"
    assert len(tool_names) >= 19


@pytest.mark.asyncio
async def test_mcp_lifecycle_and_inspection(temp_deck_path: Path):
    """Test ppt_open, ppt_inspect_presentation, ppt_inspect_slide, ppt_inspect_shape via MCP client."""
    # 1. ppt_open
    open_res = await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})
    open_data = _extract_tool_result(open_res)
    assert open_data.get("success") is True
    assert "session_id" in open_data
    assert open_data.get("slide_count") == 3

    # 2. ppt_inspect_presentation
    prs_res = await app.call_tool("ppt_inspect_presentation", {})
    prs_data = _extract_tool_result(prs_res)
    assert prs_data.get("success") is True
    assert prs_data.get("slide_count") == 3
    assert prs_data.get("width_inches") > 0

    # 3. ppt_inspect_slide
    slide_res = await app.call_tool("ppt_inspect_slide", {"slide_number": 1})
    slide_data = _extract_tool_result(slide_res)
    assert slide_data.get("success") is True
    assert slide_data.get("slide_number") == 1
    assert len(slide_data.get("shapes", [])) > 0

    first_shape_id = slide_data["shapes"][0]["shape_id"]

    # 4. ppt_inspect_shape
    shape_res = await app.call_tool("ppt_inspect_shape", {"slide_number": 1, "shape_id": first_shape_id})
    shape_data = _extract_tool_result(shape_res)
    assert shape_data.get("success") is True
    assert shape_data.get("shape", {}).get("shape_id") == first_shape_id


@pytest.mark.asyncio
async def test_mcp_editing_tools(temp_deck_path: Path):
    """Test geometry and text editing tools via MCP client."""
    # Open deck session
    await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})

    # 1. ppt_modify_shape
    mod_res = await app.call_tool(
        "ppt_modify_shape",
        {"slide_number": 1, "shape_id": 2, "dx": 0.2, "dy": 0.1},
    )
    mod_data = _extract_tool_result(mod_res)
    assert mod_data.get("success") is True

    # 2. ppt_modify_text
    txt_res = await app.call_tool(
        "ppt_modify_text",
        {"slide_number": 1, "shape_id": 2, "text": "Updated Title Text", "bold": True},
    )
    txt_data = _extract_tool_result(txt_res)
    assert txt_data.get("success") is True

    # 3. ppt_move_shape
    mv_res = await app.call_tool(
        "ppt_move_shape",
        {"slide_number": 1, "shape_id": 2, "dx": 0.1, "dy": 0.1},
    )
    mv_data = _extract_tool_result(mv_res)
    assert mv_data.get("success") is True

    # 4. ppt_resize_shape
    rz_res = await app.call_tool(
        "ppt_resize_shape",
        {"slide_number": 1, "shape_id": 2, "width": 8.0, "height": 1.5},
    )
    rz_data = _extract_tool_result(rz_res)
    assert rz_data.get("success") is True

    # 5. ppt_copy_shape
    cp_res = await app.call_tool(
        "ppt_copy_shape",
        {"slide_number": 1, "shape_id": 2, "target_slide_number": 1, "x_offset": 0.5, "y_offset": 0.5},
    )
    cp_data = _extract_tool_result(cp_res)
    assert cp_data.get("success") is True
    new_id = cp_data.get("new_shape_id")
    assert new_id is not None

    # 6. ppt_modify_ooxml
    ooxml_res = await app.call_tool(
        "ppt_modify_ooxml",
        {"slide_number": 1, "shape_id": new_id, "operation": "transparency", "transparency_percent": 30.0},
    )
    ooxml_data = _extract_tool_result(ooxml_res)
    assert ooxml_data.get("success") is True

    # 7. ppt_delete_shape
    del_res = await app.call_tool("ppt_delete_shape", {"slide_number": 1, "shape_id": new_id})
    del_data = _extract_tool_result(del_res)
    assert del_data.get("success") is True


@pytest.mark.asyncio
async def test_mcp_validation_and_comparison(temp_deck_path: Path):
    """Test ppt_validate_slide and ppt_compare_slides via MCP client."""
    await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})

    # 1. ppt_validate_slide on slide 1 (clean)
    v1_res = await app.call_tool("ppt_validate_slide", {"slide_number": 1})
    v1_data = _extract_tool_result(v1_res)
    assert v1_data.get("success") is True
    assert v1_data.get("is_valid") is True

    # 2. ppt_validate_slide on slide 3 (intentional flaws)
    v3_res = await app.call_tool("ppt_validate_slide", {"slide_number": 3})
    v3_data = _extract_tool_result(v3_res)
    assert v3_data.get("success") is True
    assert len(v3_data.get("warnings", [])) > 0

    # 3. ppt_compare_slides
    comp_res = await app.call_tool("ppt_compare_slides", {"slide_a": 1, "slide_b": 2})
    comp_data = _extract_tool_result(comp_res)
    assert comp_data.get("success") is True
    assert "geometric_match_score" in comp_data


@pytest.mark.asyncio
async def test_mcp_rendering_and_diff(temp_deck_path: Path, tmp_path: Path):
    """Test ppt_render_slide, ppt_render_presentation, ppt_visual_diff via MCP client."""
    await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})

    render_dir = tmp_path / "mcp_renders"
    render_dir.mkdir(parents=True, exist_ok=True)

    # 1. ppt_render_slide
    r1_res = await app.call_tool(
        "ppt_render_slide",
        {"slide_number": 1, "output_dir": str(render_dir), "renderer": "mock"},
    )
    r1_data = _extract_tool_result(r1_res)
    assert r1_data.get("success") is True
    img_path1 = r1_data.get("image_path")
    assert Path(img_path1).exists()

    # Render slide 2 for diffing
    r2_res = await app.call_tool(
        "ppt_render_slide",
        {"slide_number": 2, "output_dir": str(render_dir), "renderer": "mock"},
    )
    r2_data = _extract_tool_result(r2_res)
    img_path2 = r2_data.get("image_path")

    # 2. ppt_visual_diff
    diff_res = await app.call_tool(
        "ppt_visual_diff",
        {"before_image": img_path1, "after_image": img_path2},
    )
    diff_data = _extract_tool_result(diff_res)
    assert diff_data.get("success") is True
    assert "similarity_percentage" in diff_data

    # 3. ppt_render_presentation
    pres_res = await app.call_tool(
        "ppt_render_presentation",
        {"output_dir": str(render_dir), "renderer": "mock"},
    )
    pres_data = _extract_tool_result(pres_res)
    assert pres_data.get("success") is True
    assert pres_data.get("slide_count") == 3


@pytest.mark.asyncio
async def test_mcp_save_save_as_revert(temp_deck_path: Path, tmp_path: Path):
    """Test ppt_save, ppt_save_as, ppt_revert via MCP client."""
    await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})

    # Modify
    await app.call_tool("ppt_modify_shape", {"slide_number": 1, "shape_id": 2, "dx": 0.5})

    # 1. ppt_save_as
    out_copy = tmp_path / "saved_as_deck.pptx"
    sa_res = await app.call_tool("ppt_save_as", {"output_path": str(out_copy)})
    sa_data = _extract_tool_result(sa_res)
    assert sa_data.get("success") is True
    assert out_copy.exists()

    # 2. ppt_revert
    rev_res = await app.call_tool("ppt_revert", {"target": "original"})
    rev_data = _extract_tool_result(rev_res)
    assert rev_data.get("success") is True

    # 3. ppt_save
    save_res = await app.call_tool("ppt_save", {})
    save_data = _extract_tool_result(save_res)
    assert save_data.get("success") is True


@pytest.mark.asyncio
async def test_mcp_resources(temp_deck_path: Path):
    """Test MCP resources reading."""
    await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})

    # Read presentation resource
    pres_res = await app.read_resource("ppt://current/presentation")
    assert pres_res is not None
    assert len(pres_res) > 0
    pres_content = pres_res[0].content
    assert "slide_count" in pres_content

    # Read slide resource
    slide_res = await app.read_resource("ppt://current/slide/1")
    assert slide_res is not None
    assert len(slide_res) > 0
    slide_content = slide_res[0].content
    assert "shapes" in slide_content


@pytest.mark.asyncio
async def test_mcp_structured_error_handling(temp_deck_path: Path):
    """Verify that domain errors return structured JSON error payloads."""
    await app.call_tool("ppt_open", {"presentation_path": str(temp_deck_path)})

    # Non-existent shape
    bad_shape_res = await app.call_tool(
        "ppt_modify_shape",
        {"slide_number": 1, "shape_id": 9999, "dx": 0.1},
    )
    bad_shape_data = _extract_tool_result(bad_shape_res)
    assert bad_shape_data.get("success") is False
    assert bad_shape_data.get("error_type") == "ShapeNotFound"

    # Out-of-range slide number
    bad_slide_res = await app.call_tool(
        "ppt_inspect_slide",
        {"slide_number": 99},
    )
    bad_slide_data = _extract_tool_result(bad_slide_res)
    assert bad_slide_data.get("success") is False
    assert bad_slide_data.get("error_type") == "SlideNotFound"
