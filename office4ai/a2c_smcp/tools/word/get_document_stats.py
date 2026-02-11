"""word_get_document_stats MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class WordGetDocumentStatsInput(BaseModel):
    """MCP 输入模型: 获取文档统计"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")


class WordGetDocumentStatsTool(BaseTool):
    """获取 Word 文档的统计信息"""

    @property
    def name(self) -> str:
        return "word_get_document_stats"

    @property
    def description(self) -> str:
        return (
            "Get statistics for a Word document. "
            "Returns word count, character count, and paragraph count. "
            "Use this to understand the document size and content volume."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetDocumentStatsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:documentStats"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetDocumentStatsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回统计摘要 | Get tool: return stats summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        words = obs.data.get("wordCount", 0)
        chars = obs.data.get("characterCount", 0)
        paragraphs = obs.data.get("paragraphCount", 0)
        content = f"{words} words, {chars} characters, {paragraphs} paragraphs"
        return {"success": True, "content": content, "data": obs.data}
