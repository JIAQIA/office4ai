# filename: base.py
# @Time    : 2025/10/29 12:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
MCP 工具基类 | MCP Tool Base Class

定义所有工具的基础接口和通用功能
Defines base interface and common functionality for all tools
"""

from abc import ABC, abstractmethod
from typing import Any, TypeVar

from pydantic import BaseModel

from ide4ai import IDE

# 泛型类型变量 | Generic type variable
T = TypeVar("T", bound=BaseModel)


class BaseTool(ABC):
    """
    工具基类 | Tool Base Class

    所有 MCP 工具都应该继承此类并实现 execute 方法
    All MCP tools should inherit from this class and implement the execute method
    """

    def __init__(self, ide: IDE) -> None:
        """
        初始化工具 | Initialize tool

        Args:
            ide: PythonIDE 实例 | PythonIDE instance
        """
        self.ide = ide

    @property
    @abstractmethod
    def name(self) -> str:
        """
        工具名称 | Tool name

        Returns:
            str: 工具名称 | Tool name
        """
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """
        工具描述 | Tool description

        Returns:
            str: 工具描述 | Tool description
        """
        pass

    @property
    @abstractmethod
    def input_schema(self) -> dict[str, Any]:
        """
        工具输入 Schema | Tool input schema

        Returns:
            dict: JSON Schema 格式的输入定义 | Input definition in JSON Schema format
        """
        pass

    @abstractmethod
    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行工具 | Execute tool

        Args:
            arguments: 工具参数 | Tool arguments

        Returns:
            dict: 执行结果 | Execution result
        """
        pass

    def validate_input(self, arguments: dict[str, Any], model: type[T]) -> T:
        """
        验证输入参数 | Validate input arguments

        Args:
            arguments: 输入参数 | Input arguments
            model: Pydantic 模型类 | Pydantic model class

        Returns:
            T: 验证后的模型实例 | Validated model instance

        Raises:
            ValueError: 参数验证失败 | Argument validation failed
        """
        try:
            return model.model_validate(arguments)
        except Exception as e:
            raise ValueError(f"参数验证失败 | Argument validation failed: {e}") from e
