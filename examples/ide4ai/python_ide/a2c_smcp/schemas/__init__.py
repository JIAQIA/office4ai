# filename: __init__.py
# @Time    : 2025/11/03 23:40
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Python IDE MCP 工具的 Schema 定义 | Python IDE MCP Tools Schema Definitions

定义 Python IDE 特有的 MCP 工具的输入输出 Schema
Defines input/output schemas for Python IDE specific MCP tools
"""

from ide4ai.python_ide.a2c_smcp.schemas.tools import (
    NotebookEditInput,
    NotebookEditOutput,
)

__all__ = [
    "NotebookEditInput",
    "NotebookEditOutput",
]
