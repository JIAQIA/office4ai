# filename: __init__.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""MCP 资源集合 | MCP resources."""

from office4ai.a2c_smcp.resources.base import BaseResource
from office4ai.a2c_smcp.resources.window import WindowResource

__all__ = [
    "BaseResource",
    "WindowResource",
]
