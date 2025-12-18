# -*- coding: utf-8 -*-
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


class OfficeMCPServer(BaseMCPServer):
    def __init__(self, config: MCPServerConfig) -> None:
        super().__init__(config=config, server_name="office4ai-mcp")

    def _register_tools(self) -> None:
        return

    def _register_resources(self) -> None:
        return


async def async_main() -> None:
    config = MCPServerConfig()
    logger.info(
        f"启动 MCP Server | Starting MCP Server: transport={config.transport}, host={config.host}, port={config.port}",
    )

    server = OfficeMCPServer(config)
    await server.run()


def main() -> None:
    asyncio.run(async_main())
