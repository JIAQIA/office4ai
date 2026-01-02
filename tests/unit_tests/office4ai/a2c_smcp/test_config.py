# filename: test_config.py
# @Time    : 2025/12/18 19:16
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
MCPServerConfig 单元测试 | MCPServerConfig unit tests
"""

import os
from unittest.mock import patch

import pytest
from confz import DataSource

from office4ai.a2c_smcp.config import MCPServerConfig


class TestMCPServerConfig:
    """MCPServerConfig 测试类 | MCPServerConfig test class"""

    def test_default_config(self):
        """测试默认配置 | Test default configuration"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            assert config.transport == "stdio"
            assert config.host == "127.0.0.1"
            assert config.port == 8000

    def test_stdio_transport(self):
        """测试 stdio 传输模式 | Test stdio transport mode"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "stdio"})):
            config = MCPServerConfig()
            assert config.transport == "stdio"

    def test_sse_transport(self):
        """测试 sse 传输模式 | Test sse transport mode"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "sse"})):
            config = MCPServerConfig()
            assert config.transport == "sse"

    def test_streamable_http_transport(self):
        """测试 streamable-http 传输模式 | Test streamable-http transport mode"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "streamable-http"})):
            config = MCPServerConfig()
            assert config.transport == "streamable-http"

    def test_invalid_transport(self):
        """测试无效的传输模式 | Test invalid transport mode"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "invalid"})):
            with pytest.raises(ValueError):  # confz will raise validation error
                MCPServerConfig()

    def test_custom_host_and_port(self):
        """测试自定义主机和端口 | Test custom host and port"""
        with MCPServerConfig.change_config_sources(DataSource(data={"host": "0.0.0.0", "port": 9000})):
            config = MCPServerConfig()
            assert config.host == "0.0.0.0"
            assert config.port == 9000

    @pytest.mark.parametrize("port", [0, -1, 65536])
    def test_invalid_port(self, port):
        """测试无效的端口 | Test invalid port"""
        with MCPServerConfig.change_config_sources(DataSource(data={"port": port})):
            with pytest.raises(ValueError):  # confz will raise validation error
                MCPServerConfig()
