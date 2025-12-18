# filename: tools.py
# @Time    : 2025/11/03 23:40
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Python IDE MCP 工具的 Schema 定义 | Python IDE MCP Tools Schema Definitions

定义 Python IDE 特有的工具的输入输出 Schema
Defines input/output schemas for Python IDE specific tools
"""

from typing import Any, Literal

from pydantic import BaseModel, Field


class NotebookEditInput(BaseModel):
    """NotebookEdit 工具输入 Schema | NotebookEdit Tool Input Schema"""

    notebook_path: str = Field(
        ..., description="要编辑的 Jupyter notebook 文件的绝对路径（必须是绝对路径，而不是相对路径）"
    )
    cell_id: str | None = Field(
        default=None,
        description="要编辑的单元格的 ID。插入新单元格时，新单元格将插入到此 ID 的单元格之后，如果未指定则插入到开头。",
    )
    new_source: str = Field(..., description="单元格的新源代码")
    cell_type: Literal["code", "markdown"] | None = Field(
        default=None,
        description="单元格类型（code 或 markdown）。如果未指定，默认为当前单元格类型。如果使用 edit_mode=insert，则此项为必填。",
    )
    edit_mode: Literal["replace", "insert", "delete"] = Field(
        default="replace",
        description="编辑类型（replace、insert、delete）。默认为 replace。",
    )


class NotebookEditOutput(BaseModel):
    """NotebookEdit 工具输出 Schema | NotebookEdit Tool Output Schema"""

    success: bool = Field(..., description="是否成功执行")
    message: str = Field(default="", description="操作结果消息")
    error: str | None = Field(default=None, description="错误信息(如果有)")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据",
    )
