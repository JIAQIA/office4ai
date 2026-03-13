"""ppt_insert_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import SlideTableInsertOptions


class PptInsertTableInput(BaseModel):
    """MCP 输入模型: 插入表格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    options: SlideTableInsertOptions = Field(..., description="Table insertion options (rows, columns, data)")


class PptInsertTableTool(BaseTool):
    """在幻灯片上插入表格"""

    @property
    def name(self) -> str:
        return "ppt_insert_table"

    @property
    def description(self) -> str:
        return (
            "Insert a table on a PowerPoint slide. "
            "Requires specifying the number of rows and columns. "
            "Supports optional initial data and position settings."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptInsertTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "insert:table"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptInsertTableInput
