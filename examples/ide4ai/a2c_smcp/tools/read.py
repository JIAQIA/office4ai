# filename: read.py
# @Time    : 2025/11/03 17:42
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Read 工具实现 | Read Tool Implementation

提供在 IDE 环境中读取文件内容的能力
Provides the ability to read file contents in the IDE environment
"""

from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.schemas import ReadInput, ReadOutput
from ide4ai.a2c_smcp.tools.base import BaseTool
from ide4ai.environment.workspace.schema import Position, Range


class ReadTool(BaseTool):
    """
    Read 文件读取工具 | Read File Reading Tool

    通过 PythonIDE 的工作区环境读取文件内容
    Reads file contents through PythonIDE's workspace environment
    """

    @property
    def name(self) -> str:
        return "Read"

    @property
    def description(self) -> str:
        return """从本地文件系统读取文件。你可以使用此工具直接访问任何文件。
假设此工具能够读取机器上的所有文件。如果用户提供了文件路径，请假定该路径有效。即使文件不存在也可以读取；将返回错误。

使用说明：
- file_path 参数必须是绝对路径，而不是相对路径
- 默认情况下，从文件开头读取最多 2000 行
- 你可以选择指定行偏移量和限制（对于长文件特别方便），但建议通过不提供这些参数来读取整个文件
- 任何超过 2000 个字符的行都将被截断
- 结果使用 cat -n 格式返回，行号从 1 开始
- 此工具允许读取图像（例如 PNG、JPG 等）。读取图像文件时，内容会以视觉方式呈现
- 此工具可以读取 PDF 文件（.pdf）。PDF 逐页处理，提取文本和视觉内容以供分析
- 此工具可以读取 Jupyter notebooks（.ipynb 文件）并返回所有单元格及其输出，结合代码、文本和可视化
- 此工具只能读取文件，不能读取目录。要读取目录，请通过 Bash 工具使用 ls 命令
- 你可以在单个响应中调用多个工具。最好同时推测性地读取多个可能有用的文件
- 你会经常被要求读取屏幕截图。如果用户提供了屏幕截图的路径，请始终使用此工具查看该路径的文件。此工具适用于所有临时文件路径
- 如果你读取的文件存在但内容为空，你将收到系统提醒警告而不是文件内容"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return ReadInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 Read 读取 | Execute Read operation

        Args:
            arguments: 包含文件路径和读取选项的参数字典 | Arguments dict containing file path and read options

        Returns:
            dict: 执行结果，包含文件内容 | Execution result with file content
        """
        # 验证输入参数 | Validate input arguments
        try:
            read_input = self.validate_input(arguments, ReadInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return ReadOutput(
                success=False,
                content="",
                error=err_info,
            ).model_dump()

        logger.info(
            f"执行 Read 读取 | Executing Read: file_path={read_input.file_path}, offset={read_input.offset}, limit={read_input.limit}",
        )

        try:
            # 检查 workspace 是否可用 | Check if workspace is available
            if self.ide.workspace is None:
                raise RuntimeError("Workspace 未初始化 | Workspace is not initialized")

            # 构建文件 URI | Build file URI
            file_path = read_input.file_path
            if not file_path.startswith("file://"):
                file_uri = f"file://{file_path}"
            else:
                file_uri = file_path

            # 构建 Range 参数（如果提供了 offset 和 limit）| Build Range parameter (if offset and limit provided)
            code_range = None
            if read_input.offset is not None:
                # 计算结束行号 | Calculate end line number
                # 如果提供了 limit，则读取 offset 到 offset+limit-1 行
                # If limit is provided, read from offset to offset+limit-1
                # 如果没有提供 limit，则读取到文件末尾（使用一个很大的数字）
                # If limit is not provided, read to end of file (use a large number)
                start_line = read_input.offset
                end_line = start_line + read_input.limit - 1 if read_input.limit else 999999

                code_range = Range(
                    start_position=Position(line=start_line, character=1),
                    end_position=Position(line=end_line, character=1),
                )

            # 调用 workspace 的 read_file 方法 | Call workspace's read_file method
            content = self.ide.workspace.read_file(
                uri=file_uri,
                with_line_num=True,  # 始终显示行号 | Always show line numbers
                code_range=code_range,
            )

            # 构造输出 | Construct output
            output = ReadOutput(
                success=True,
                content=content,
                metadata={
                    "file_path": file_path,
                    "offset": read_input.offset,
                    "limit": read_input.limit,
                },
            )

            logger.info(
                f"Read 读取完成 | Read completed: file_path={file_path}, content_length={len(content)}",
            )

            return output.model_dump()

        except FileNotFoundError as e:
            # 处理文件不存在错误 | Handle file not found error
            logger.error(f"文件不存在 | File not found: {e}")
            error_output = ReadOutput(
                success=False,
                content="",
                error=f"文件不存在 | File not found: {read_input.file_path}",
            )
            return error_output.model_dump()

        except ValueError as e:
            # 处理路径验证错误 | Handle path validation errors
            logger.error(f"路径验证失败 | Path validation failed: {e}")
            error_output = ReadOutput(
                success=False,
                content="",
                error=str(e),
            )
            return error_output.model_dump()

        except Exception as e:
            logger.exception(f"执行 Read 读取时发生错误 | Error executing Read: {e}")

            # 返回错误结果 | Return error result
            error_output = ReadOutput(
                success=False,
                content="",
                error=str(e),
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
