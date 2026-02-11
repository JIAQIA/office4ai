# filename: conftest.py
# @Time    : 2025/12/18 19:19
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
集成测试配置 | Integration test configuration
"""

from pathlib import Path

import pytest

# 项目根目录路径 | Project root path
project_root = Path(__file__).parent.parent.parent


@pytest.fixture(scope="session")
def project_root_path():
    """项目根目录路径 fixture | Project root path fixture"""
    return project_root


@pytest.fixture(scope="session")
def mcp_server_params(project_root_path):
    """MCP Server 参数 fixture | MCP Server parameters fixture"""
    from mcp.client.stdio import StdioServerParameters

    return StdioServerParameters(
        command="uv",
        args=["run", "python", "-m", "office4ai.office.mcp.server"],
        cwd=str(project_root_path),
    )
