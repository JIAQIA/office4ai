# filename: glob.py
# @Time    : 2025/11/01 17:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Glob 工具实现 | Glob Tool Implementation

提供在 IDE 环境中使用通配符模式匹配文件的能力
Provides the ability to match files using glob patterns in the IDE environment
"""

from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.schemas import GlobInput, GlobOutput
from ide4ai.a2c_smcp.tools.base import BaseTool


class GlobTool(BaseTool):
    """
    Glob 文件匹配工具 | Glob File Matching Tool

    通过 PythonIDE 的工作区环境使用通配符模式匹配文件
    Matches files using glob patterns through PythonIDE's workspace environment
    """

    @property
    def name(self) -> str:
        return "Glob"

    @property
    def description(self) -> str:
        return """快速的文件模式匹配工具，适用于任何代码库大小
- 支持 glob 模式，如 "**/*.js" 或 "src/**/*.ts"
- 返回按修改时间排序的匹配文件路径
- 当您需要按名称模式查找文件时使用此工具
- 您可以在单个响应中调用多个工具。如果多个搜索可能有用，最好同时执行它们"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return GlobInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 Glob 文件匹配 | Execute Glob file matching

        Args:
            arguments: 包含模式和路径的参数字典 | Arguments dict containing pattern and path

        Returns:
            dict: 执行结果，包含匹配的文件列表 | Execution result with matched files list
        """
        # 验证输入参数 | Validate input arguments
        try:
            glob_input = self.validate_input(arguments, GlobInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return GlobOutput(
                success=False,
                files=[],
                error=err_info,
            ).model_dump()

        logger.info(f"执行 Glob 匹配 | Executing Glob matching: pattern={glob_input.pattern}, path={glob_input.path}")

        try:
            # 检查 workspace 是否可用 | Check if workspace is available
            if self.ide.workspace is None:
                raise RuntimeError("Workspace 未初始化 | Workspace is not initialized")

            # 调用 workspace 的 glob_files 方法
            # Call workspace's glob_files method
            matched_files = self.ide.workspace.glob_files(
                pattern=glob_input.pattern,
                path=glob_input.path,
            )

            # 构造输出 | Construct output
            output = GlobOutput(
                success=True,
                files=matched_files,
                metadata={
                    "pattern": glob_input.pattern,
                    "path": glob_input.path,
                    "count": len(matched_files),
                },
            )

            logger.info(f"Glob 匹配完成 | Glob matching completed: found {len(matched_files)} files")

            return output.model_dump()

        except ValueError as e:
            # 处理路径验证错误 | Handle path validation errors
            logger.error(f"路径验证失败 | Path validation failed: {e}")
            error_output = GlobOutput(
                success=False,
                files=[],
                error=str(e),
            )
            return error_output.model_dump()

        except Exception as e:
            logger.exception(f"执行 Glob 匹配时发生错误 | Error executing Glob matching: {e}")

            # 返回错误结果 | Return error result
            error_output = GlobOutput(
                success=False,
                files=[],
                error=str(e),
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
