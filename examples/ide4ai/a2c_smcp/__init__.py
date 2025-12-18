# filename: __init__.py
# @Time    : 2025/11/03 20:25
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
MCP (Model Context Protocol) 基础模块 | MCP Base Module

提供通用的 MCP Server 实现，支持多语言 IDE 扩展
Provides generic MCP Server implementation, supporting multi-language IDE extensions
"""

from .config import MCPServerConfig
from .server import BaseMCPServer

__all__ = [
    "BaseMCPServer",
    "MCPServerConfig",
]
