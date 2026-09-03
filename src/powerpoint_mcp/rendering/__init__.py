from powerpoint_mcp.rendering.com_lifecycle import (
    cleanup_mcp_com_processes,
    com_powerpoint_session,
    defensive_file_operation,
    get_active_mcp_pids,
    get_powerpoint_pids,
    register_mcp_pid,
    terminate_mcp_pid,
    unregister_mcp_pid,
)
from powerpoint_mcp.rendering.image_diff import (
    VisualDiffResult,
    visual_diff,
)
from powerpoint_mcp.rendering.renderer import (
    BaseRenderer,
    LibreOfficeRenderer,
    NullRenderer,
    PowerPointRenderer,
    get_available_renderer,
)
from powerpoint_mcp.rendering.visual_compare import (
    SlideComparisonResult,
    compare_slides,
)

__all__ = [
    "BaseRenderer",
    "LibreOfficeRenderer",
    "NullRenderer",
    "PowerPointRenderer",
    "SlideComparisonResult",
    "VisualDiffResult",
    "cleanup_mcp_com_processes",
    "com_powerpoint_session",
    "compare_slides",
    "defensive_file_operation",
    "get_active_mcp_pids",
    "get_available_renderer",
    "get_powerpoint_pids",
    "register_mcp_pid",
    "terminate_mcp_pid",
    "unregister_mcp_pid",
    "visual_diff",
]
