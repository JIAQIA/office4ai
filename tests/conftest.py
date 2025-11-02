"""
Pytest configuration and fixtures for Office4AI tests.
"""

import os
import tempfile
from collections.abc import Generator
from pathlib import Path

import pytest


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Create a temporary directory for tests."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_docx_path(temp_dir: Path) -> Path:
    """Create a sample docx file path."""
    return temp_dir / "sample.docx"


@pytest.fixture
def sample_xlsx_path(temp_dir: Path) -> Path:
    """Create a sample xlsx file path."""
    return temp_dir / "sample.xlsx"


@pytest.fixture
def sample_pptx_path(temp_dir: Path) -> Path:
    """Create a sample pptx file path."""
    return temp_dir / "sample.pptx"


@pytest.fixture
def libreoffice_path() -> str:
    """Get LibreOffice path from environment or use default."""
    default_paths = {
        "darwin": "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "linux": "/usr/bin/soffice",
        "win32": "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    }

    import sys

    platform = sys.platform

    if platform.startswith("darwin"):
        default = default_paths["darwin"]
    elif platform.startswith("linux"):
        default = default_paths["linux"]
    elif platform.startswith("win"):
        default = default_paths["win32"]
    else:
        default = "soffice"

    return os.environ.get("LIBREOFFICE_PATH", default)
