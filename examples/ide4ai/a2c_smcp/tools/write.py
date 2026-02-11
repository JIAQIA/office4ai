# filename: write.py
# @Time    : 2025/11/03 23:40
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Write 工具实现 | Write Tool Implementation

提供在 IDE 环境中写入文件的能力
Provides the ability to write files in the IDE environment
"""

from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.schemas import WriteInput, WriteOutput
from ide4ai.a2c_smcp.tools.base import BaseTool


class WriteTool(BaseTool):
    """
    Write 文件写入工具 | Write File Writing Tool

    通过 PythonIDE 的工作区环境写入文件内容
    Writes file contents through PythonIDE's workspace environment
    """

    @property
    def name(self) -> str:
        return "Write"

    @property
    def description(self) -> str:
        return """将文件写入本地文件系统。

使用说明：
- 如果提供的路径存在文件，此工具将覆盖现有文件。
- 如果这是现有文件，你必须先使用 Read 工具读取该文件。如果你没有先读取文件，此工具将失败。
- 始终优先编辑代码库中的现有文件。除非明确要求，否则切勿写入新文件。
- 切勿主动创建文档文件（*.md）或 README 文件。仅在用户明确请求时创建文档文件。
- 仅在用户明确请求时使用表情符号。除非被要求，否则避免向文件写入表情符号。"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return WriteInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 Write 写入 | Execute Write operation

        Args:
            arguments: 包含文件路径和内容的参数字典 | Arguments dict containing file path and content

        Returns:
            dict: 执行结果 | Execution result
        """
        # 验证输入参数 | Validate input arguments
        try:
            write_input = self.validate_input(arguments, WriteInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return WriteOutput(
                success=False,
                message="",
                error=err_info,
            ).model_dump()

        logger.info(
            f"执行 Write 写入 | Executing Write: file_path={write_input.file_path}",
        )

        try:
            # 检查 workspace 是否可用 | Check if workspace is available
            if self.ide.workspace is None:
                raise RuntimeError("Workspace 未初始化 | Workspace is not initialized")

            # 构建文件 URI | Build file URI
            file_path = write_input.file_path
            if not file_path.startswith("file://"):
                file_uri = f"file://{file_path}"
            else:
                file_uri = file_path

            # 检查文件是否存在 | Check if file exists
            import os

            file_exists = os.path.exists(file_path)

            if file_exists:
                # 如果文件存在，打开文件并替换全部内容 | If file exists, open and replace all content
                text_model = self.ide.workspace.open_file(uri=file_uri)

                # 获取当前文件的全部内容范围 | Get the full content range
                from ide4ai.environment.workspace.schema import Position, Range, SingleEditOperation

                line_count = text_model.get_line_count()
                last_line_length = len(text_model.get_line_content(line_count))

                # 创建一个编辑操作，替换整个文件内容 | Create an edit operation to replace entire file content
                full_range = Range(
                    start_position=Position(line=1, character=1),
                    end_position=Position(line=line_count, character=last_line_length + 1),
                )

                edit_operation = SingleEditOperation(
                    range=full_range,
                    text=write_input.content,
                )

                # 应用编辑 | Apply edit
                undo_edits, diagnostics = self.ide.workspace.apply_edit(
                    uri=file_uri,
                    edits=[edit_operation],
                    compute_undo_edits=True,
                )

                # 保存文件 | Save file
                self.ide.workspace.save_file(uri=file_uri)

                message = f"成功写入文件: {file_path}"

                # 添加诊断信息 | Add diagnostics info
                if diagnostics:
                    from ide4ai.python_ide.workspace import PyWorkspace

                    message += "\n" + PyWorkspace._format_diagnostics(diagnostics)

                output = WriteOutput(
                    success=True,
                    message=message,
                    metadata={
                        "file_path": file_path,
                        "file_existed": True,
                        "content_length": len(write_input.content),
                    },
                )

            else:
                # 如果文件不存在，创建新文件 | If file doesn't exist, create new file
                text_model, diagnostics = self.ide.workspace.create_file(uri=file_uri)

                if text_model is None:
                    raise RuntimeError(f"无法创建文件 | Cannot create file: {file_path}")

                # 如果创建的文件有初始内容，需要先清空 | If created file has initial content, clear it first
                if text_model.get_value():
                    from ide4ai.environment.workspace.schema import Position, Range, SingleEditOperation

                    line_count = text_model.get_line_count()
                    last_line_length = len(text_model.get_line_content(line_count))

                    full_range = Range(
                        start_position=Position(line=1, character=1),
                        end_position=Position(line=line_count, character=last_line_length + 1),
                    )

                    clear_operation = SingleEditOperation(
                        range=full_range,
                        text="",
                    )

                    self.ide.workspace.apply_edit(
                        uri=file_uri,
                        edits=[clear_operation],
                        compute_undo_edits=False,
                    )

                # 写入内容 | Write content
                from ide4ai.environment.workspace.schema import Position, Range, SingleEditOperation

                insert_operation = SingleEditOperation(
                    range=Range(
                        start_position=Position(line=1, character=1),
                        end_position=Position(line=1, character=1),
                    ),
                    text=write_input.content,
                )

                undo_edits, diagnostics = self.ide.workspace.apply_edit(
                    uri=file_uri,
                    edits=[insert_operation],
                    compute_undo_edits=True,
                )

                # 保存文件 | Save file
                self.ide.workspace.save_file(uri=file_uri)

                message = f"成功创建并写入文件: {file_path}"

                # 添加诊断信息 | Add diagnostics info
                if diagnostics:
                    from ide4ai.python_ide.workspace import PyWorkspace

                    message += "\n" + PyWorkspace._format_diagnostics(diagnostics)

                output = WriteOutput(
                    success=True,
                    message=message,
                    metadata={
                        "file_path": file_path,
                        "file_existed": False,
                        "content_length": len(write_input.content),
                    },
                )

            logger.info(
                f"Write 写入完成 | Write completed: file_path={file_path}",
            )

            return output.model_dump()

        except FileNotFoundError as e:
            # 处理文件不存在错误 | Handle file not found error
            logger.error(f"文件操作失败 | File operation failed: {e}")
            error_output = WriteOutput(
                success=False,
                message="",
                error=f"文件操作失败 | File operation failed: {write_input.file_path}",
            )
            return error_output.model_dump()

        except ValueError as e:
            # 处理路径验证错误 | Handle path validation errors
            logger.error(f"路径验证失败 | Path validation failed: {e}")
            error_output = WriteOutput(
                success=False,
                message="",
                error=str(e),
            )
            return error_output.model_dump()

        except Exception as e:
            logger.exception(f"执行 Write 写入时发生错误 | Error executing Write: {e}")

            # 返回错误结果 | Return error result
            error_output = WriteOutput(
                success=False,
                message="",
                error=str(e),
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
