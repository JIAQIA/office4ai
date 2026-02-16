# filename: test_mcp_protocol.py
# @Time    : 2025/12/18 19:18
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
MCP 协议集成测试 | MCP protocol integration tests
"""

from pathlib import Path

import pytest
from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client
from pydantic import AnyUrl

# 项目根目录路径 | Project root path
project_root = Path(__file__).parent.parent.parent.parent.parent


@pytest.mark.integration
class TestMCPProtocol:
    """MCP 协议测试类 | MCP protocol test class"""

    async def test_mcp_handshake(self):
        """测试 MCP 握手 | Test MCP handshake"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                # 初始化会话 | Initialize session
                result = await session.initialize()

                # 验证服务器信息 | Verify server info
                assert result is not None
                assert result.serverInfo.name == "office4ai"

    async def test_list_tools(self):
        """测试 list_tools 返回已注册工具 | Test list_tools returns registered tools"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 获取工具列表 | Get tools list
                tools_result = await session.list_tools()
                assert len(tools_result.tools) == 42  # 21 Word + 21 PPT

                # 验证工具名称前缀 | Verify tool name prefix
                tool_names = {t.name for t in tools_result.tools}
                assert all(name.startswith(("word_", "ppt_")) for name in tool_names)

    async def test_list_resources(self):
        """测试 list_resources 返回已注册资源 | Test list_resources returns registered resources"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 获取资源列表 | Get resources list
                resources_result = await session.list_resources()
                assert len(resources_result.resources) == 2  # ConnectedDocuments + WindowResource

                # 验证资源 URI | Verify resource URIs
                resource_uris = {str(r.uri) for r in resources_result.resources}
                assert "office://workspace/documents" in resource_uris
                assert any(uri.startswith("window://office4ai") for uri in resource_uris)

    async def test_call_tool_not_found(self):
        """测试调用不存在的工具 | Test calling non-existent tool"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 尝试调用不存在的工具 | Try to call non-existent tool
                result = await session.call_tool("non_existent_tool", {})
                assert len(result.content) == 1
                assert "未找到工具 | Tool not found" in result.content[0].text

    async def test_read_resource_not_found(self):
        """测试读取不存在的资源 | Test reading non-existent resource"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 尝试读取不存在的资源 | Try to read non-existent resource
                from mcp.shared.exceptions import McpError

                with pytest.raises(McpError, match="未找到资源 | Resource not found"):
                    await session.read_resource(AnyUrl("office://non/existent/resource"))
