"""word_export_content MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.word import ExportOptions


class WordExportContentInput(BaseModel):
    """MCP 输入模型: 导出文档内容"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    format: Literal["text", "html", "markdown"] = Field(
        ...,
        description="Export format: text (plain text), html, or markdown",
    )
    options: ExportOptions | None = Field(
        None,
        description="Export options (includeImages, includeTables)",
    )


class WordExportContentTool(BaseTool):
    """导出 Word 文档内容为指定格式"""

    @property
    def name(self) -> str:
        return "word_export_content"

    @property
    def description(self) -> str:
        return (
            "Export the content of a Word document in a specified format. "
            "Supports plain text, HTML, and Markdown output formats. "
            "Use this to get the full document content for processing or analysis."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordExportContentInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "export:content"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordExportContentInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """导出类工具: 直接返回导出内容 | Export tool: return exported content directly"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        content = obs.data.get("content", "")
        return {"success": True, "content": content, "data": obs.data}
