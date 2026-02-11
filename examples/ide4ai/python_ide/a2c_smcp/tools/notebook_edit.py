# filename: notebook_edit.py
# @Time    : 2025/11/03 23:40
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
NotebookEdit 工具实现 | NotebookEdit Tool Implementation

提供在 IDE 环境中编辑 Jupyter Notebook 的能力
Provides the ability to edit Jupyter Notebooks in the IDE environment
"""

import json
from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.tools.base import BaseTool
from ide4ai.python_ide.a2c_smcp.schemas import NotebookEditInput, NotebookEditOutput


class NotebookEditTool(BaseTool):
    """
    NotebookEdit Jupyter Notebook 编辑工具 | NotebookEdit Jupyter Notebook Editing Tool

    通过 PythonIDE 的工作区环境编辑 Jupyter Notebook 文件
    Edits Jupyter Notebook files through PythonIDE's workspace environment
    """

    @property
    def name(self) -> str:
        return "NotebookEdit"

    @property
    def description(self) -> str:
        return """完全替换 Jupyter notebook（.ipynb 文件）中特定单元格的内容。

Jupyter notebooks 是结合代码、文本和可视化的交互式文档，通常用于数据分析和科学计算。

使用说明：
- notebook_path 参数必须是绝对路径，而不是相对路径
- cell_id 是单元格的 ID（0 索引）
- 使用 edit_mode=insert 在 cell_id 指定的索引处添加新单元格
- 使用 edit_mode=delete 删除 cell_id 指定索引处的单元格"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return NotebookEditInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 NotebookEdit 编辑 | Execute NotebookEdit operation

        Args:
            arguments: 包含 notebook 路径和编辑参数的参数字典 | Arguments dict containing notebook path and edit parameters

        Returns:
            dict: 执行结果 | Execution result
        """
        # 验证输入参数 | Validate input arguments
        try:
            edit_input = self.validate_input(arguments, NotebookEditInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return NotebookEditOutput(
                success=False,
                message="",
                error=err_info,
            ).model_dump()

        logger.info(
            f"执行 NotebookEdit 编辑 | Executing NotebookEdit: notebook_path={edit_input.notebook_path}, "
            f"cell_id={edit_input.cell_id}, edit_mode={edit_input.edit_mode}",
        )

        try:
            # 检查 workspace 是否可用 | Check if workspace is available
            if self.ide.workspace is None:
                raise RuntimeError("Workspace 未初始化 | Workspace is not initialized")

            # 获取文件路径 | Get file path
            notebook_path = edit_input.notebook_path
            # 如果是 file:// URI，提取实际路径 | If it's a file:// URI, extract actual path
            if notebook_path.startswith("file://"):
                notebook_path = notebook_path[7:]  # 移除 "file://" 前缀 | Remove "file://" prefix

            # 检查文件是否存在 | Check if file exists
            import os

            if not os.path.exists(notebook_path):
                raise FileNotFoundError(f"Notebook 文件不存在 | Notebook file not found: {notebook_path}")

            # 检查文件扩展名 | Check file extension
            if not notebook_path.endswith(".ipynb"):
                raise ValueError(f"文件必须是 .ipynb 格式 | File must be .ipynb format: {notebook_path}")

            # 读取 notebook 文件 | Read notebook file
            with open(notebook_path, encoding="utf-8") as f:
                notebook_data = json.load(f)

            # 验证 notebook 结构 | Validate notebook structure
            if "cells" not in notebook_data:
                raise ValueError(
                    "无效的 notebook 格式：缺少 'cells' 字段 | Invalid notebook format: missing 'cells' field"
                )

            cells = notebook_data["cells"]

            if edit_input.edit_mode == "insert":
                # 插入新单元格 | Insert new cell
                if edit_input.cell_type is None:
                    raise ValueError("插入模式需要指定 cell_type | cell_type is required for insert mode")

                new_cell: dict[str, Any] = {
                    "cell_type": edit_input.cell_type,
                    "metadata": {},
                    "source": edit_input.new_source.splitlines(keepends=True) if edit_input.new_source else [],
                }

                # 如果是 code 类型，添加必要的字段 | Add necessary fields for code type
                if edit_input.cell_type == "code":
                    new_cell["execution_count"] = None
                    new_cell["outputs"] = []

                # 确定插入位置 | Determine insert position
                if edit_input.cell_id is None:
                    # 插入到开头 | Insert at beginning
                    insert_index = 0
                else:
                    # 查找指定 cell_id 的单元格 | Find cell with specified cell_id
                    insert_index = None
                    for idx, cell in enumerate(cells):
                        if cell.get("id") == edit_input.cell_id:
                            insert_index = idx + 1  # 插入到该单元格之后 | Insert after this cell
                            break

                    if insert_index is None:
                        raise ValueError(
                            f"未找到 cell_id: {edit_input.cell_id} | Cell with id {edit_input.cell_id} not found"
                        )

                cells.insert(insert_index, new_cell)
                message = f"成功在位置 {insert_index} 插入新单元格"

            elif edit_input.edit_mode == "delete":
                # 删除单元格 | Delete cell
                if edit_input.cell_id is None:
                    raise ValueError("删除模式需要指定 cell_id | cell_id is required for delete mode")

                # 查找并删除指定 cell_id 的单元格 | Find and delete cell with specified cell_id
                delete_index = None
                for idx, cell in enumerate(cells):
                    if cell.get("id") == edit_input.cell_id:
                        delete_index = idx
                        break

                if delete_index is None:
                    raise ValueError(
                        f"未找到 cell_id: {edit_input.cell_id} | Cell with id {edit_input.cell_id} not found"
                    )

                cells.pop(delete_index)
                message = f"成功删除位置 {delete_index} 的单元格"

            else:  # replace
                # 替换单元格内容 | Replace cell content
                if edit_input.cell_id is None:
                    raise ValueError("替换模式需要指定 cell_id | cell_id is required for replace mode")

                # 查找指定 cell_id 的单元格 | Find cell with specified cell_id
                target_cell = None
                cell_index = None
                for idx, cell in enumerate(cells):
                    if cell.get("id") == edit_input.cell_id:
                        target_cell = cell
                        cell_index = idx
                        break

                if target_cell is None:
                    raise ValueError(
                        f"未找到 cell_id: {edit_input.cell_id} | Cell with id {edit_input.cell_id} not found"
                    )

                # 更新单元格内容 | Update cell content
                target_cell["source"] = edit_input.new_source.splitlines(keepends=True) if edit_input.new_source else []

                # 如果指定了 cell_type，更新单元格类型 | Update cell type if specified
                if edit_input.cell_type is not None:
                    old_type = target_cell["cell_type"]
                    target_cell["cell_type"] = edit_input.cell_type

                    # 如果从 markdown 改为 code，添加必要字段 | Add necessary fields when changing from markdown to code
                    if old_type == "markdown" and edit_input.cell_type == "code":
                        target_cell["execution_count"] = None
                        target_cell["outputs"] = []
                    # 如果从 code 改为 markdown，删除 code 特有字段 | Remove code-specific fields when changing to markdown
                    elif old_type == "code" and edit_input.cell_type == "markdown":
                        target_cell.pop("execution_count", None)
                        target_cell.pop("outputs", None)

                message = f"成功替换位置 {cell_index} 的单元格内容"

            # 写回 notebook 文件 | Write back to notebook file
            with open(notebook_path, "w", encoding="utf-8") as f:
                json.dump(notebook_data, f, indent=1, ensure_ascii=False)
                f.write("\n")  # 添加文件末尾的换行符 | Add trailing newline

            output = NotebookEditOutput(
                success=True,
                message=message,
                metadata={
                    "notebook_path": notebook_path,
                    "edit_mode": edit_input.edit_mode,
                    "cell_id": edit_input.cell_id,
                    "cell_type": edit_input.cell_type,
                    "total_cells": len(cells),
                },
            )

            logger.info(
                f"NotebookEdit 编辑完成 | NotebookEdit completed: notebook_path={notebook_path}",
            )

            return output.model_dump()

        except FileNotFoundError as e:
            # 处理文件不存在错误 | Handle file not found error
            logger.error(f"文件不存在 | File not found: {e}")
            error_output = NotebookEditOutput(
                success=False,
                message="",
                error=f"文件不存在 | File not found: {edit_input.notebook_path}",
            )
            return error_output.model_dump()

        except json.JSONDecodeError as e:
            # 处理 JSON 解析错误 | Handle JSON decode errors
            logger.error(f"Notebook 文件格式错误 | Invalid notebook format: {e}")
            error_output = NotebookEditOutput(
                success=False,
                message="",
                error=f"Notebook 文件格式错误 | Invalid notebook format: {e}",
            )
            return error_output.model_dump()

        except ValueError as e:
            # 处理验证错误 | Handle validation errors
            logger.error(f"验证失败 | Validation failed: {e}")
            error_output = NotebookEditOutput(
                success=False,
                message="",
                error=str(e),
            )
            return error_output.model_dump()

        except Exception as e:
            logger.exception(f"执行 NotebookEdit 编辑时发生错误 | Error executing NotebookEdit: {e}")

            # 返回错误结果 | Return error result
            error_output = NotebookEditOutput(
                success=False,
                message="",
                error=str(e),
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
