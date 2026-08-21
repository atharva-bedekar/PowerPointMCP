"""Path resolution, session workspace management, and backup naming utilities."""

from datetime import datetime, timezone
import os
from pathlib import Path
import shutil
import time
from typing import Dict, Optional, Union

DEFAULT_WORKSPACE_DIR_NAME = ".ppt-agent"
DEFAULT_SESSIONS_DIR_NAME = "sessions"
DEFAULT_WORKING_FILENAME = "working.pptx"
DEFAULT_ORIGINAL_FILENAME = "original.pptx"
DEFAULT_METADATA_FILENAME = "metadata.json"
DEFAULT_BACKUPS_DIR_NAME = "backups"
DEFAULT_RENDERS_DIR_NAME = "renders"
DEFAULT_DIFFS_DIR_NAME = "diffs"


def get_workspace_dir(base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return the absolute path to the .ppt-agent workspace root directory.

    Resolution order:
    1. PPT_WORKSPACE_DIR environment variable (if set).
    2. Explicit base_dir / .ppt-agent (if base_dir is provided).
    3. Path.cwd() / .ppt-agent (default).
    """
    env_dir = os.environ.get("PPT_WORKSPACE_DIR")
    if env_dir:
        return Path(env_dir).resolve()
    if base_dir is not None:
        return (Path(base_dir) / DEFAULT_WORKSPACE_DIR_NAME).resolve()
    return (Path.cwd() / DEFAULT_WORKSPACE_DIR_NAME).resolve()


def get_sessions_root_dir(base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return the path to .ppt-agent/sessions/ directory."""
    return get_workspace_dir(base_dir) / DEFAULT_SESSIONS_DIR_NAME


def get_session_dir(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return the path to a specific session workspace: .ppt-agent/sessions/<session_id>/."""
    return get_sessions_root_dir(base_dir) / session_id


def get_session_working_path(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return path to session's working.pptx."""
    return get_session_dir(session_id, base_dir) / DEFAULT_WORKING_FILENAME


def get_session_original_path(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return path to session's original.pptx."""
    return get_session_dir(session_id, base_dir) / DEFAULT_ORIGINAL_FILENAME


def get_session_metadata_path(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return path to session's metadata.json."""
    return get_session_dir(session_id, base_dir) / DEFAULT_METADATA_FILENAME


def get_session_backups_dir(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return path to session's backups/ directory."""
    return get_session_dir(session_id, base_dir) / DEFAULT_BACKUPS_DIR_NAME


def get_session_renders_dir(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return path to session's renders/ directory."""
    return get_session_dir(session_id, base_dir) / DEFAULT_RENDERS_DIR_NAME


def get_session_diffs_dir(session_id: str, base_dir: Optional[Union[str, Path]] = None) -> Path:
    """Return path to session's diffs/ directory."""
    return get_session_dir(session_id, base_dir) / DEFAULT_DIFFS_DIR_NAME


def generate_backup_filename(
    base_name: str = "presentation",
    timestamp: Optional[datetime] = None,
) -> str:
    """Generate a standard timestamped backup filename following convention:

    presentation.backup-YYYYMMDD-HHMMSS.pptx

    Args:
        base_name: Prefix for backup file (e.g. 'presentation' or presentation stem).
        timestamp: Optional specific datetime (defaults to now).

    Returns:
        Formatted backup filename string.
    """
    if base_name.lower().endswith(".pptx"):
        base_name = base_name[:-5]

    if timestamp is None:
        # Use local or UTC now
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    else:
        ts = timestamp.strftime("%Y%m%d-%H%M%S")

    return f"{base_name}.backup-{ts}.pptx"


def ensure_session_dirs(
    session_id: str,
    base_dir: Optional[Union[str, Path]] = None,
) -> Dict[str, Path]:
    """Create all required directory structures for an isolated session workspace.

    Args:
        session_id: UUID or identifier for the session.
        base_dir: Optional workspace base directory override.

    Returns:
        Dictionary mapping path keys to resolved Path objects.
    """
    s_dir = get_session_dir(session_id, base_dir)
    backups_dir = get_session_backups_dir(session_id, base_dir)
    renders_dir = get_session_renders_dir(session_id, base_dir)
    diffs_dir = get_session_diffs_dir(session_id, base_dir)

    s_dir.mkdir(parents=True, exist_ok=True)
    backups_dir.mkdir(parents=True, exist_ok=True)
    renders_dir.mkdir(parents=True, exist_ok=True)
    diffs_dir.mkdir(parents=True, exist_ok=True)

    return {
        "session_dir": s_dir,
        "working_path": get_session_working_path(session_id, base_dir),
        "original_path": get_session_original_path(session_id, base_dir),
        "metadata_path": get_session_metadata_path(session_id, base_dir),
        "backups_dir": backups_dir,
        "renders_dir": renders_dir,
        "diffs_dir": diffs_dir,
    }


def cleanup_session(
    session_id: str,
    base_dir: Optional[Union[str, Path]] = None,
) -> bool:
    """Remove a session directory and all its contents cleanly.

    Args:
        session_id: Session identifier to remove.
        base_dir: Optional workspace base directory override.

    Returns:
        True if the directory was deleted or did not exist.
    """
    s_dir = get_session_dir(session_id, base_dir)
    if s_dir.exists():
        shutil.rmtree(s_dir, ignore_errors=True)
        return not s_dir.exists()
    return True


def cleanup_old_sessions(
    max_age_hours: int = 24,
    base_dir: Optional[Union[str, Path]] = None,
) -> int:
    """Scan and delete session directories older than max_age_hours.

    Args:
        max_age_hours: Maximum allowable age in hours before removal (default 24).
        base_dir: Optional workspace base directory override.

    Returns:
        Count of deleted sessions.
    """
    sessions_root = get_sessions_root_dir(base_dir)
    if not sessions_root.exists():
        return 0

    now = time.time()
    max_age_sec = max_age_hours * 3600
    cleaned_count = 0

    for session_entry in sessions_root.iterdir():
        if session_entry.is_dir():
            try:
                mtime = session_entry.stat().st_mtime
                if (now - mtime) > max_age_sec:
                    shutil.rmtree(session_entry, ignore_errors=True)
                    cleaned_count += 1
            except Exception:
                continue

    return cleaned_count
