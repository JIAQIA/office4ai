# filename: __init__.py
# @Time    : 2025/10/29 12:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Python IDE MCP Server 封装 | Python IDE MCP Server Wrapper

将 PythonIDE 的能力封装为 MCP Server 对外提供服务
Wraps PythonIDE capabilities as MCP Server for external services
"""

from ide4ai.a2c_smcp.config import MCPServerConfig
from ide4ai.python_ide.a2c_smcp.server import PythonIDEMCPServer

__all__ = ["MCPServerConfig", "PythonIDEMCPServer"]
