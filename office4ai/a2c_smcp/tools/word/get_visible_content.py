"""word_get_visible_content MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.word import GetContentOptions


class WordGetVisibleContentInput(BaseModel):
    """MCP 输入模型: 获取可见内容"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    options: GetContentOptions | None = Field(None, description="Content retrieval options")


class WordGetVisibleContentTool(BaseTool):
    """获取 Word 文档中当前可见的内容"""

    @property
    def name(self) -> str:
        return "word_get_visible_content"

    @property
    def description(self) -> str:
        return (
            "Get the currently visible content in a Word document viewport. "
            "Returns the visible text, elements (text/images/tables), and metadata. "
            "Use this to understand the document context the user is currently viewing."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetVisibleContentInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:visibleContent"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetVisibleContentInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回适配文本 | Get tool: return adapted text"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        content = obs.data.get("text", "")
        return {"success": True, "content": content, "data": obs.data}
