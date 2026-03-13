"""
Office4AI 测试的 Pytest 配置和 fixtures | Pytest configuration and fixtures for Office4AI tests
"""

import os

# Prevent tests from creating a logs/ directory in the working directory
os.environ.setdefault("OFFICE4AI_LOG_DIR", "")
import tempfile
from collections.abc import Generator
from pathlib import Path

import pytest


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """为测试创建临时目录 | Create a temporary directory for tests."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_docx_path(temp_dir: Path) -> Path:
    """创建示例 docx 文件路径 | Create a sample docx file path."""
    return temp_dir / "sample.docx"


@pytest.fixture
def sample_xlsx_path(temp_dir: Path) -> Path:
    """创建示例 xlsx 文件路径 | Create a sample xlsx file path."""
    return temp_dir / "sample.xlsx"


@pytest.fixture
def sample_pptx_path(temp_dir: Path) -> Path:
    """创建示例 pptx 文件路径 | Create a sample pptx file path."""
    return temp_dir / "sample.pptx"


@pytest.fixture
def libreoffice_path() -> str:
    """从环境变量获取 LibreOffice 路径或使用默认值 | Get LibreOffice path from environment or use default."""
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
