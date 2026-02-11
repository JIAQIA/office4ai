"""word_get_styles MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.word import GetStylesOptions


class WordGetStylesInput(BaseModel):
    """MCP 输入模型: 获取样式列表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    options: GetStylesOptions | None = Field(
        None,
        description="Style retrieval options (includeBuiltIn, includeCustom, includeUnused, detailedInfo)",
    )


class WordGetStylesTool(BaseTool):
    """获取 Word 文档中可用的样式"""

    @property
    def name(self) -> str:
        return "word_get_styles"

    @property
    def description(self) -> str:
        return (
            "Get available styles in a Word document. "
            "Returns a list of paragraph, character, table, and list styles. "
            "Supports filtering by built-in/custom/unused styles and optional detailed info."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetStylesInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:styles"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetStylesInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回样式名列表 | Get tool: return style name list"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        styles = obs.data.get("styles", [])
        style_names = [s.get("name", "Unknown") for s in styles] if isinstance(styles, list) else []
        content = ", ".join(style_names) if style_names else "No styles found"
        return {"success": True, "content": content, "data": obs.data}
