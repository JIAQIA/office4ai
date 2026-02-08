# filename: server.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

import asyncio

from loguru import logger

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.server import BaseMCPServer
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


class OfficeMCPServer(BaseMCPServer):
    """
    Office 级别的统一 MCP Server | Office-level unified MCP Server

    一个 Server 实例同时处理 Word、PPT、Excel 三种文档类型。
    Socket.IO Server 的生命周期与 MCP Server 一致。
    """

    def __init__(self, config: MCPServerConfig) -> None:
        # 同步：创建 workspace 实例 (未启动)
        self.workspace = OfficeWorkspace(
            host=config.host,
            port=config.socketio_port,
        )
        super().__init__(config=config, server_name="office4ai")

    def _register_tools(self) -> None:
        """注册所有平台的工具 | Register all platform tools"""
        from office4ai.a2c_smcp.tools.word import (
            WordAppendTextTool,
            WordGetSelectedContentTool,
            WordGetVisibleContentTool,
            WordInsertEquationTool,
            WordInsertImageTool,
            WordInsertTableTool,
            WordInsertTextTool,
            WordInsertTOCTool,
            WordReplaceTextTool,
        )

        word_tools = [
            WordGetSelectedContentTool(self.workspace),
            WordGetVisibleContentTool(self.workspace),
            WordInsertTextTool(self.workspace),
            WordAppendTextTool(self.workspace),
            WordReplaceTextTool(self.workspace),
            WordInsertImageTool(self.workspace),
            WordInsertTableTool(self.workspace),
            WordInsertEquationTool(self.workspace),
            WordInsertTOCTool(self.workspace),
        ]

        for tool in word_tools:
            self.tools[tool.name] = tool

        logger.info(f"已注册 {len(word_tools)} 个 Word 工具 | Registered {len(word_tools)} Word tools")

        # PPT 工具 (未来) | PPT tools (future)
        # Excel 工具 (未来) | Excel tools (future)

    def _register_resources(self) -> None:
        """注册资源 | Register resources"""
        from office4ai.a2c_smcp.resources.connected_documents import ConnectedDocumentsResource

        docs_resource = ConnectedDocumentsResource(self.workspace)
        self.resources[docs_resource.base_uri] = docs_resource

    async def _async_startup(self) -> None:
        """启动 OfficeWorkspace (Socket.IO Server) | Start OfficeWorkspace"""
        logger.info("启动 OfficeWorkspace | Starting OfficeWorkspace...")
        await self.workspace.start()

    async def _async_shutdown(self) -> None:
        """停止 OfficeWorkspace (Socket.IO Server) | Stop OfficeWorkspace"""
        logger.info("停止 OfficeWorkspace | Stopping OfficeWorkspace...")
        await self.workspace.stop()


async def async_main() -> None:
    config = MCPServerConfig()
    logger.info(
        f"启动 MCP Server | Starting MCP Server: transport={config.transport}, host={config.host}, port={config.port}",
    )

    server = OfficeMCPServer(config)
    await server.run()


def main() -> None:
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
