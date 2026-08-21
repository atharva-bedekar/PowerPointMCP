"""Utility functions and modules for PowerPoint MCP server."""

from powerpoint_mcp.utils.logging import get_logger
from powerpoint_mcp.utils.paths import (
    cleanup_old_sessions,
    cleanup_session,
    ensure_session_dirs,
    generate_backup_filename,
    get_session_backups_dir,
    get_session_diffs_dir,
    get_session_dir,
    get_session_metadata_path,
    get_session_original_path,
    get_session_renders_dir,
    get_session_working_path,
    get_workspace_dir,
)
from powerpoint_mcp.utils.validation import (
    IssueSeverity,
    PresentationValidationResult,
    SlideIssue,
    SlideValidationResult,
    validate_presentation,
    validate_slide,
)

__all__ = [
    "IssueSeverity",
    "PresentationValidationResult",
    "SlideIssue",
    "SlideValidationResult",
    "cleanup_old_sessions",
    "cleanup_session",
    "ensure_session_dirs",
    "generate_backup_filename",
    "get_logger",
    "get_session_backups_dir",
    "get_session_diffs_dir",
    "get_session_dir",
    "get_session_metadata_path",
    "get_session_original_path",
    "get_session_renders_dir",
    "get_session_working_path",
    "get_workspace_dir",
    "validate_presentation",
    "validate_slide",
]
