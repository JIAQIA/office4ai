# filename: test_server.py
# @Time    : 2025/12/18 19:15
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
OfficeMCPServer 单元测试 | OfficeMCPServer unit tests
"""

import os
from unittest.mock import patch

import pytest
from confz import DataSource

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.office.mcp.server import OfficeMCPServer


class TestOfficeMCPServer:
    """OfficeMCPServer 测试类 | OfficeMCPServer test class"""

    def test_server_initialization(self):
        """测试服务器初始化 | Test server initialization"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert server.config == config
            assert server.server.name == "office4ai-mcp"
            assert len(server.tools) == 0
            assert len(server.resources) == 0

    @pytest.mark.asyncio
    async def test_register_tools(self):
        """测试工具注册 | Test tool registration"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            # Milestone 0: 应该没有注册任何工具 | Milestone 0: No tools should be registered
            assert len(server.tools) == 0

    @pytest.mark.asyncio
    async def test_register_resources(self):
        """测试资源注册 | Test resource registration"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            # Milestone 0: 应该没有注册任何资源 | Milestone 0: No resources should be registered
            assert len(server.resources) == 0

    @pytest.mark.parametrize("transport", ["stdio", "sse", "streamable-http"])
    def test_server_supports_all_transports(self, transport):
        """测试服务器支持所有传输模式 | Test server supports all transport modes"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": transport})):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert server.config.transport == transport

    def test_server_name(self):
        """测试服务器名称 | Test server name"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert server.server.name == "office4ai-mcp"
