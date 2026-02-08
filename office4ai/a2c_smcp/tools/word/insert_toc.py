"""word_insert_toc MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import TOCOptions


class WordInsertTOCInput(BaseModel):
    """MCP 输入模型: 插入目录"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    options: TOCOptions | None = Field(None, description="TOC options (maxLevel, heading styles to include)")


class WordInsertTOCTool(BaseTool):
    """在 Word 文档中插入目录"""

    @property
    def name(self) -> str:
        return "word_insert_toc"

    @property
    def description(self) -> str:
        return (
            "Insert a table of contents (TOC) into a Word document. "
            "By default includes headings up to level 3. "
            "You can customize the maximum heading level and which styles to include."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertTOCInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:toc"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertTOCInput
