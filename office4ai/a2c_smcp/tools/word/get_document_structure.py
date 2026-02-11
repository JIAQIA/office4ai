"""word_get_document_structure MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class WordGetDocumentStructureInput(BaseModel):
    """MCP 输入模型: 获取文档结构"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")


class WordGetDocumentStructureTool(BaseTool):
    """获取 Word 文档的结构信息"""

    @property
    def name(self) -> str:
        return "word_get_document_structure"

    @property
    def description(self) -> str:
        return (
            "Get the structural overview of a Word document. "
            "Returns counts of sections, paragraphs, tables, and images. "
            "Use this to understand the document layout before making modifications."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetDocumentStructureInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:documentStructure"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetDocumentStructureInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回结构摘要 | Get tool: return structure summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        sections = obs.data.get("sectionCount", 0)
        paragraphs = obs.data.get("paragraphCount", 0)
        tables = obs.data.get("tableCount", 0)
        images = obs.data.get("imageCount", 0)
        content = f"{sections} sections, {paragraphs} paragraphs, {tables} tables, {images} images"
        return {"success": True, "content": content, "data": obs.data}
