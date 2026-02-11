# filename: grep.py
# @Time    : 2025/11/01 20:47
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Grep 工具实现 | Grep Tool Implementation

提供在 IDE 环境中使用 ripgrep 搜索文件内容的能力
Provides the ability to search file contents using ripgrep in the IDE environment
"""

from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.schemas import GrepInput, GrepOutput
from ide4ai.a2c_smcp.tools.base import BaseTool


class GrepTool(BaseTool):
    """
    Grep 文件内容搜索工具 | Grep File Content Search Tool

    通过 PythonIDE 的工作区环境使用 ripgrep 搜索文件内容
    Searches file contents using ripgrep through PythonIDE's workspace environment
    """

    @property
    def name(self) -> str:
        return "Grep"

    @property
    def description(self) -> str:
        return """一个基于 ripgrep 构建的强大搜索工具

使用说明：
- 对于搜索任务，请始终使用 Grep 工具。切勿使用 Bash 命令 grep 或 rg。Grep 工具已针对正确的权限和访问进行了优化。
- 支持完整的正则表达式语法（例如："log.*Error", "function\\s+\\w+"）
- 使用 glob 参数（例如："*.js", "**/*.tsx"）或 type 参数（例如："js", "py", "rust"）来过滤文件
- 输出模式："content"显示匹配的行，"files_with_matches"仅显示文件路径（默认），"count"显示匹配计数
- 对于需要多轮进行的开放式搜索，请使用 Task 工具
- 模式语法：使用 ripgrep（而非 grep）- 字面量花括号需要转义（例如，在 Go 代码中查找 interface{} 需使用 interface\\{\\}）
- 多行匹配：默认情况下，模式仅在单行内匹配。对于跨行模式（如 struct \\{[\\s\\S]*?field），请使用 multiline: true"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return GrepInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 Grep 搜索 | Execute Grep search

        Args:
            arguments: 包含搜索模式和选项的参数字典 | Arguments dict containing search pattern and options

        Returns:
            dict: 执行结果，包含搜索输出 | Execution result with search output
        """
        # 验证输入参数 | Validate input arguments
        try:
            grep_input = self.validate_input(arguments, GrepInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return GrepOutput(
                success=False,
                output="",
                matched=False,
                error=err_info,
            ).model_dump()

        logger.info(
            f"执行 Grep 搜索 | Executing Grep search: pattern={grep_input.pattern}, path={grep_input.path}, "
            f"output_mode={grep_input.output_mode}",
        )

        try:
            # 检查 workspace 是否可用 | Check if workspace is available
            if self.ide.workspace is None:
                raise RuntimeError("Workspace 未初始化 | Workspace is not initialized")

            # 调用 workspace 的 grep_files 方法
            # Call workspace's grep_files method
            result = self.ide.workspace.grep_files(
                pattern=grep_input.pattern,
                path=grep_input.path,
                glob=grep_input.glob,
                output_mode=grep_input.output_mode,
                context_before=grep_input.context_before,
                context_after=grep_input.context_after,
                context=grep_input.context,
                line_number=grep_input.line_number,
                case_insensitive=grep_input.case_insensitive,
                file_type=grep_input.file_type,
                head_limit=grep_input.head_limit,
                multiline=grep_input.multiline,
            )

            # 构造输出 | Construct output
            output = GrepOutput(
                success=result["success"],
                output=result["output"],
                matched=result["matched"],
                metadata=result["metadata"],
            )

            logger.info(
                f"Grep 搜索完成 | Grep search completed: matched={output.matched}, output_length={len(output.output)}",
            )

            return output.model_dump()

        except ValueError as e:
            # 处理路径验证错误 | Handle path validation errors
            logger.error(f"路径验证失败 | Path validation failed: {e}")
            error_output = GrepOutput(
                success=False,
                output="",
                matched=False,
                error=str(e),
            )
            return error_output.model_dump()

        except RuntimeError as e:
            # 处理 ripgrep 执行错误 | Handle ripgrep execution errors
            logger.error(f"Ripgrep 执行失败 | Ripgrep execution failed: {e}")
            error_output = GrepOutput(
                success=False,
                output="",
                matched=False,
                error=str(e),
            )
            return error_output.model_dump()

        except Exception as e:
            logger.exception(f"执行 Grep 搜索时发生错误 | Error executing Grep search: {e}")

            # 返回错误结果 | Return error result
            error_output = GrepOutput(
                success=False,
                output="",
                matched=False,
                error=str(e),
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
