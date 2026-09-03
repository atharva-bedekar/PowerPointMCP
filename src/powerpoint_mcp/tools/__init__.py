"""MCP tool modules, inspection, editing, rendering, and session versioning controllers."""

from powerpoint_mcp.tools.editing import (
    ppt_add_picture,
    ppt_batch_modify_shapes,
    ppt_batch_modify_table_cells,
    ppt_batch_modify_tables,
    ppt_batch_modify_text,
    ppt_copy_shape,
    ppt_delete_shape,
    ppt_merge_table_cells,
    ppt_modify_ooxml,
    ppt_modify_shape,
    ppt_modify_text,
    ppt_move_shape,
    ppt_replace_picture,
    ppt_resize_shape,
    ppt_set_table_geometry,
    ppt_style_table,
)
from powerpoint_mcp.tools.inspection import (
    handle_tool_errors,
    ppt_compare_slides,
    ppt_inspect_presentation,
    ppt_inspect_shape,
    ppt_inspect_slide,
    ppt_inspect_table,
    ppt_inspect_text,
    ppt_validate_slide,
    ppt_validate_slides,
)
from powerpoint_mcp.tools.rendering import (
    ppt_render_presentation,
    ppt_render_slide,
    ppt_render_slides,
    ppt_visual_diff,
)
from powerpoint_mcp.tools.versioning import (
    Session,
    SessionManager,
    create_backup,
    get_current_session,
    get_session,
    get_session_manager,
    open_presentation,
    revert_session,
    save_as,
    save_session,
)

__all__ = [
    # Session / Versioning
    "Session",
    "SessionManager",
    "create_backup",
    "get_current_session",
    "get_session",
    "get_session_manager",
    "open_presentation",
    "revert_session",
    "save_as",
    "save_session",
    # Inspection
    "handle_tool_errors",
    "ppt_compare_slides",
    "ppt_inspect_presentation",
    "ppt_inspect_shape",
    "ppt_inspect_slide",
    "ppt_inspect_table",
    "ppt_inspect_text",
    "ppt_validate_slide",
    "ppt_validate_slides",
    # Editing
    "ppt_add_picture",
    "ppt_batch_modify_shapes",
    "ppt_batch_modify_table_cells",
    "ppt_batch_modify_tables",
    "ppt_batch_modify_text",
    "ppt_copy_shape",
    "ppt_delete_shape",
    "ppt_merge_table_cells",
    "ppt_modify_ooxml",
    "ppt_modify_shape",
    "ppt_modify_text",
    "ppt_move_shape",
    "ppt_replace_picture",
    "ppt_resize_shape",
    "ppt_set_table_geometry",
    "ppt_style_table",
    # Rendering
    "ppt_render_presentation",
    "ppt_render_slide",
    "ppt_render_slides",
    "ppt_visual_diff",
]


