#!/usr/bin/env python
# filename: start_mcp_server.py
# @Time    : 2025/12/18 19:05
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
MCP Server 启动脚本 | MCP Server startup script
用于测试 Office4AI MCP Server 的基础功能 | Used to test basic functionality of Office4AI MCP Server
"""

import asyncio

from office4ai.office.mcp.server import async_main


if __name__ == "__main__":
    # 设置日志级别 | Set log level
    import logging

    logging.basicConfig(level=logging.INFO)

    print("启动 Office4AI MCP Server... | Starting Office4AI MCP Server...")
    asyncio.run(async_main())
