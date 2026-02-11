# filename: test_server_base.py
# @Time    : 2025/12/18 19:17
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
BaseMCPServer 单元测试 | BaseMCPServer unit tests
"""

import os
from unittest.mock import patch

import pytest
from confz import DataSource

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.server import BaseMCPServer


class ConcreteMCPServer(BaseMCPServer):
    """用于测试的具体 MCP Server 实现 | Concrete MCP Server implementation for testing"""

    def _register_tools(self):
        pass

    def _register_resources(self):
        pass


class TestBaseMCPServer:
    """BaseMCPServer 测试类 | BaseMCPServer test class"""

    def test_initialization(self):
        """测试初始化 | Test initialization"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            assert server.config == config
            assert server.server.name == "test-server"
            assert isinstance(server.tools, dict)
            assert isinstance(server.resources, dict)

    def test_register_tools_method(self):
        """测试注册工具方法 | Test register tools method"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            # 应该调用 _register_tools | Should call _register_tools
            assert hasattr(server, "_register_tools")

    def test_register_resources_method(self):
        """测试注册资源方法 | Test register resources method"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            # 应该调用 _register_resources | Should call _register_resources
            assert hasattr(server, "_register_resources")

    @pytest.mark.asyncio
    async def test_run_stdio_transport(self):
        """测试 stdio 传输模式运行 | Test running with stdio transport"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "stdio"})):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            # stdio 需要实际的输入输出流，在单元测试中应该跳过或模拟 | stdio needs actual streams, should skip or mock in unit tests
            with patch("office4ai.a2c_smcp.server.stdio_server") as mock_stdio:
                mock_stdio.side_effect = AttributeError("stdio needs actual streams")
                with pytest.raises(AttributeError):
                    await server.run()

    @pytest.mark.asyncio
    async def test_run_invalid_transport(self):
        """测试无效传输模式 | Test invalid transport mode"""
        # 创建一个模拟的无效配置 | Create a mock invalid config
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "invalid"})):
            with pytest.raises(ValueError):  # confz will raise validation error
                MCPServerConfig()
