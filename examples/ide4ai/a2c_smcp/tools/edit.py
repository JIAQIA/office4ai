# filename: edit.py
# @Time    : 2025/11/03 18:05
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Edit 工具实现 | Edit Tool Implementation

提供在 IDE 环境中执行精确字符串替换的能力
Provides the ability to perform exact string replacements in the IDE environment
"""

from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.schemas import EditInput, EditOutput
from ide4ai.a2c_smcp.tools.base import BaseTool


class EditTool(BaseTool):
    """
    Edit 文件编辑工具 | Edit File Editing Tool

    通过 PythonIDE 的工作区环境执行精确字符串替换
    Performs exact string replacements through PythonIDE's workspace environment
    """

    @property
    def name(self) -> str:
        return "Edit"

    @property
    def description(self) -> str:
        return """在文件中执行精确的字符串替换。

使用说明：
- 在编辑之前，你必须在对话中至少使用一次 `Read` 工具。如果你在没有读取文件的情况下尝试编辑，此工具将报错。
- 从 Read 工具输出编辑文本时，请确保保留行号前缀之后显示的确切缩进（制表符/空格）。行号前缀格式为：空格 + 行号 + 制表符。该制表符之后的所有\
内容都是要匹配的实际文件内容。切勿在 old_string 或 new_string 中包含行号前缀的任何部分。
- 始终优先编辑代码库中的现有文件。除非明确要求，否则切勿编写新文件。
- 仅在用户明确请求时使用表情符号。除非被要求，否则避免向文件添加表情符号。
- 如果 `old_string` 在文件中不唯一，编辑将失败。提供一个包含更多周围上下文的较大字符串以使其唯一，或使用 `replace_all` 更改 \
`old_string` 的每个实例。
- 使用 `replace_all` 在整个文件中替换和重命名字符串。如果你想重命名变量，此参数很有用。
- 如果 `old_string` 和 `new_string` 相同，编辑将失败。这被视为无操作并将引发错误。"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return EditInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 Edit 编辑 | Execute Edit operation

        Args:
            arguments: 包含文件路径和替换参数的参数字典 | Arguments dict containing file path and replacement parameters

        Returns:
            dict: 执行结果，包含替换信息 | Execution result with replacement information
        """
        # 验证输入参数 | Validate input arguments
        try:
            edit_input = self.validate_input(arguments, EditInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return EditOutput(
                success=False,
                message="",
                replacements_made=0,
                error=err_info,
            ).model_dump()

        logger.info(
            f"执行 Edit 编辑 | Executing Edit: file_path={edit_input.file_path}, replace_all={edit_input.replace_all}",
        )

        # 验证 old_string 和 new_string 不相同 | Validate old_string and new_string are different
        if edit_input.old_string == edit_input.new_string:
            err_info = "old_string 和 new_string 不能相同 | old_string and new_string cannot be identical"
            logger.error(err_info)
            return EditOutput(
                success=False,
                message="",
                replacements_made=0,
                error=err_info,
            ).model_dump()

        try:
            # 检查 workspace 是否可用 | Check if workspace is available
            if self.ide.workspace is None:
                raise RuntimeError("Workspace 未初始化 | Workspace is not initialized")

            # 构建文件 URI | Build file URI
            file_path = edit_input.file_path
            if not file_path.startswith("file://"):
                file_uri = f"file://{file_path}"
            else:
                file_uri = file_path

            # 首先读取文件以确保文件存在 | First read file to ensure it exists
            try:
                file_content = self.ide.workspace.read_file(uri=file_uri, with_line_num=False)
            except Exception as e:
                raise FileNotFoundError(f"无法读取文件 | Cannot read file: {file_path}") from e

            # 检查 old_string 是否存在于文件中 | Check if old_string exists in file
            if edit_input.old_string not in file_content:
                err_info = "在文件中未找到 old_string | old_string not found in file"
                logger.error(err_info)
                return EditOutput(
                    success=False,
                    message="",
                    replacements_made=0,
                    error=err_info,
                ).model_dump()

            # 检查 old_string 的出现次数 | Check occurrences of old_string
            occurrences = file_content.count(edit_input.old_string)

            if occurrences == 0:
                err_info = "在文件中未找到 old_string"
                logger.error(err_info)
                return EditOutput(
                    success=False,
                    message="",
                    replacements_made=0,
                    error=err_info,
                ).model_dump()

            # 如果有多个匹配且未设置 replace_all，则报错 | Error if multiple matches without replace_all
            if occurrences > 1 and not edit_input.replace_all:
                err_info = f"old_string 在文件中不唯一（找到 {occurrences} 个匹配项）。请提供更大的上下文字符串使其唯一，或设置 replace_all=true"
                logger.error(err_info)
                return EditOutput(
                    success=False,
                    message="",
                    replacements_made=0,
                    error=err_info,
                ).model_dump()

            # 使用 replace_in_file 方法执行替换 | Use replace_in_file method to perform replacement
            # 注意：这里使用精确字符串匹配，不使用正则表达式 | Note: Use exact string matching, not regex
            undo_edits, diagnostics = self.ide.workspace.replace_in_file(
                uri=file_uri,
                query=edit_input.old_string,
                replacement=edit_input.new_string,
                is_regex=False,  # 精确字符串匹配 | Exact string matching
                match_case=True,  # 区分大小写 | Case sensitive
                compute_undo_edits=True,  # 计算撤销编辑 | Compute undo edits
            )

            # 计算实际替换次数 | Calculate actual replacements made
            replacements_made = len(undo_edits) if undo_edits else 0

            # 构造成功消息 | Construct success message
            message = f"成功替换 {replacements_made} 处"

            # 添加诊断信息 | Add diagnostics info
            if diagnostics:
                from ide4ai.python_ide.workspace import PyWorkspace

                message += "\n" + PyWorkspace._format_diagnostics(diagnostics)

            # 构造输出 | Construct output
            output = EditOutput(
                success=True,
                message=message,
                replacements_made=replacements_made,
                metadata={
                    "file_path": file_path,
                    "replace_all": edit_input.replace_all,
                    "undo_edits": [str(edit) for edit in undo_edits] if undo_edits else [],
                },
            )

            logger.info(
                f"Edit 编辑完成: file_path={file_path}, replacements_made={replacements_made}",
            )

            return output.model_dump()

        except FileNotFoundError as e:
            # 处理文件不存在错误 | Handle file not found error
            logger.error(f"文件不存在: {e}")
            error_output = EditOutput(
                success=False,
                message="",
                replacements_made=0,
                error=f"文件不存在: {edit_input.file_path}",
            )
            return error_output.model_dump()

        except ValueError as e:
            # 处理路径验证错误 | Handle path validation errors
            logger.error(f"路径验证失败: {e}")
            error_output = EditOutput(
                success=False,
                message="",
                replacements_made=0,
                error=str(e),
            )
            return error_output.model_dump()

        except Exception as e:
            logger.exception(f"执行 Edit 编辑时发生错误: {e}")

            # 返回错误结果 | Return error result
            error_output = EditOutput(
                success=False,
                message="",
                replacements_made=0,
                error=str(e),
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
