"""word_get_selected_content MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.word import GetContentOptions


class WordGetSelectedContentInput(BaseModel):
    """MCP 输入模型: 获取选中内容"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    options: GetContentOptions | None = Field(None, description="Content retrieval options")


class WordGetSelectedContentTool(BaseTool):
    """获取 Word 文档中的选中内容"""

    @property
    def name(self) -> str:
        return "word_get_selected_content"

    @property
    def description(self) -> str:
        return (
            "Get the currently selected content in a Word document. "
            "Returns the selected text, elements (text/images/tables), and metadata. "
            "Use this when you need to understand what the user has highlighted."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetSelectedContentInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:selectedContent"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetSelectedContentInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回适配文本 | Get tool: return adapted text"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        content = obs.data.get("text", "")
        return {"success": True, "content": content, "data": obs.data}
