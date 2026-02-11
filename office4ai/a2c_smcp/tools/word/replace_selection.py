"""word_replace_selection MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import ReplaceContent


class WordReplaceSelectionInput(BaseModel):
    """MCP 输入模型: 替换选中内容"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    content: ReplaceContent = Field(
        ...,
        description="Replacement content: provide text and/or images with optional formatting",
    )


class WordReplaceSelectionTool(BaseTool):
    """替换 Word 文档中当前选中的内容"""

    @property
    def name(self) -> str:
        return "word_replace_selection"

    @property
    def description(self) -> str:
        return (
            "Replace the currently selected content in a Word document. "
            "Accepts replacement text with optional formatting, or images. "
            "The current selection will be replaced with the provided content."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordReplaceSelectionInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "replace:selection"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordReplaceSelectionInput
