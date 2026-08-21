"""Pytest fixtures for PowerPoint MCP test suite.

Provides fixtures for:
- Synthetic 3-slide presentation path and in-memory Presentation objects
- Temporary isolated workspaces and disposable copy presentations
- Environment variable isolation and renderer capability detection
- Helper fixtures for in-memory testing and assertions
"""

import os
import shutil
import sys
from pathlib import Path
from typing import Generator

import pytest
from pptx import Presentation

from tests.fixtures.create_synthetic_deck import create_synthetic_deck


@pytest.fixture(scope="session")
def project_root() -> Path:
    """Return the absolute Path to the project root directory."""
    return Path(__file__).resolve().parent.parent


@pytest.fixture(scope="session")
def fixtures_dir(project_root: Path) -> Path:
    """Return the Path to the tests/fixtures directory."""
    fixtures_path = project_root / "tests" / "fixtures"
    fixtures_path.mkdir(parents=True, exist_ok=True)
    return fixtures_path


@pytest.fixture(scope="session")
def synthetic_deck_path(fixtures_dir: Path) -> Path:
    """Ensure tests/fixtures/synthetic_sample.pptx exists and return its path.

    This presentation contains 3 slides:
    - Slide 1: Quarterly Performance Overview (KPI cards, dashboard chart)
    - Slide 2: Operational Architecture (2-column layout, process flow diagram, footer)
    - Slide 3: Audit & Compliance Issues (overlapping shapes, boundary clipping, tiny fonts, text overflow)
    """
    deck_path = fixtures_dir / "synthetic_sample.pptx"
    if not deck_path.exists():
        create_synthetic_deck(deck_path, force=True)
    return deck_path


@pytest.fixture
def sample_presentation(synthetic_deck_path: Path) -> Presentation:
    """Return a fresh in-memory python-pptx Presentation instance of the synthetic deck."""
    return Presentation(str(synthetic_deck_path))


@pytest.fixture
def synthetic_deck_bytes(synthetic_deck_path: Path) -> bytes:
    """Return raw bytes of the synthetic deck for binary transfer testing."""
    return synthetic_deck_path.read_bytes()


@pytest.fixture
def temp_workspace_dir(tmp_path: Path) -> Path:
    """Create and return an isolated temporary workspace directory."""
    workspace = tmp_path / "workspace"
    workspace.mkdir(parents=True, exist_ok=True)
    agent_dir = workspace / ".ppt-agent"
    (agent_dir / "sessions").mkdir(parents=True, exist_ok=True)
    (agent_dir / "backups").mkdir(parents=True, exist_ok=True)
    (agent_dir / "renders").mkdir(parents=True, exist_ok=True)
    (agent_dir / "diffs").mkdir(parents=True, exist_ok=True)
    return workspace


@pytest.fixture
def temp_deck_path(synthetic_deck_path: Path, tmp_path: Path) -> Path:
    """Create a temporary disposable copy of the synthetic presentation."""
    copy_path = tmp_path / "working_sample.pptx"
    shutil.copy2(synthetic_deck_path, copy_path)
    return copy_path


@pytest.fixture
def temp_presentation(temp_deck_path: Path) -> Presentation:
    """Return a Presentation instance pointing to a disposable temporary copy."""
    return Presentation(str(temp_deck_path))


@pytest.fixture
def clean_ppt_env(temp_workspace_dir: Path) -> Generator[None, None, None]:
    """Isolate environment variables during test execution."""
    old_env = os.environ.copy()
    os.environ["PPT_WORKSPACE_DIR"] = str(temp_workspace_dir / ".ppt-agent")
    os.environ["PPT_BACKUP_ENABLED"] = "true"
    os.environ["PPT_DEFAULT_OUTPUT_DIR"] = str(temp_workspace_dir / "output")
    try:
        yield
    finally:
        os.environ.clear()
        os.environ.update(old_env)


@pytest.fixture(scope="session")
def has_powerpoint_com() -> bool:
    """Detect if Microsoft PowerPoint COM automation is available on this machine."""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"):
            return True
    except Exception:
        pass
    try:
        import win32com.client
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:
        return False


@pytest.fixture(scope="session")
def has_libreoffice() -> bool:
    """Detect if LibreOffice headless binary is available on the system PATH."""
    soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
    return soffice_path is not None
