# filename: base.py
# @Time    : 2025/11/04 16:48
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
MCP 资源基类 | MCP Resource Base Class

定义所有资源的基础接口和通用功能
Defines base interface and common functionality for all resources
"""

from abc import ABC, abstractmethod
from typing import Any
from urllib.parse import urlparse

from ide4ai import IDE


class BaseResource(ABC):
    """
    资源基类 | Resource Base Class

    所有 MCP 资源都应该继承此类并实现相关方法
    All MCP resources should inherit from this class and implement related methods
    """

    def __init__(self, ide: IDE) -> None:
        """
        初始化资源 | Initialize resource

        Args:
            ide: IDE 实例 | IDE instance
        """
        self.ide = ide

    @property
    @abstractmethod
    def uri(self) -> str:
        """
        资源完整 URI（包含查询参数）| Full resource URI (with query parameters)

        Returns:
            str: 资源的完整 URI | Full URI of the resource
        """
        pass

    @property
    def base_uri(self) -> str:
        """
        资源基础 URI（不含查询参数）| Base resource URI (without query parameters)

        用作资源字典的 key，确保相同资源不同参数使用同一个实例
        Used as key in resource dict to ensure same resource with different params uses same instance

        Returns:
            str: 不含查询参数的基础 URI | Base URI without query parameters
        """
        parsed = urlparse(self.uri)
        # 返回 scheme://netloc/path，不包含查询参数
        # Return scheme://netloc/path without query parameters
        return f"{parsed.scheme}://{parsed.netloc}{parsed.path}"

    @property
    @abstractmethod
    def name(self) -> str:
        """
        资源名称 | Resource name

        Returns:
            str: 资源名称 | Resource name
        """
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """
        资源描述 | Resource description

        Returns:
            str: 资源描述 | Resource description
        """
        pass

    @property
    @abstractmethod
    def mime_type(self) -> str:
        """
        资源 MIME 类型 | Resource MIME type

        Returns:
            str: MIME 类型，如 "text/plain", "application/json" 等
                MIME type, e.g., "text/plain", "application/json", etc.
        """
        pass

    @abstractmethod
    async def read(self) -> str:
        """
        读取资源内容 | Read resource content

        Returns:
            str: 资源内容 | Resource content
        """
        pass

    @abstractmethod
    def update_from_uri(self, uri: str) -> None:
        """
        从 URI 更新资源参数 | Update resource parameters from URI

        当用户使用不同的查询参数访问资源时调用此方法更新资源状态
        Called when user accesses resource with different query parameters

        Args:
            uri: 包含新参数的完整 URI | Full URI with new parameters
        """
        # 默认实现：不做任何操作，子类可以重写此方法
        # Default implementation: do nothing, subclasses can override
        pass

    def to_dict(self) -> dict[str, Any]:
        """
        转换为字典格式 | Convert to dictionary format

        Returns:
            dict: 资源的字典表示 | Dictionary representation of the resource
        """
        return {
            "uri": self.uri,
            "name": self.name,
            "description": self.description,
            "mimeType": self.mime_type,
        }
