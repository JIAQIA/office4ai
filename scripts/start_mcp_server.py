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
import sys
from pathlib import Path

from office4ai.office.mcp.server import async_main

# 添加项目根目录到 Python 路径 | Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))


if __name__ == "__main__":
    # 设置日志级别 | Set log level
    import logging

    logging.basicConfig(level=logging.INFO)

    print("启动 Office4AI MCP Server... | Starting Office4AI MCP Server...")
    asyncio.run(async_main())
