"""word_replace_text MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import ReplaceOptions


class WordReplaceTextInput(BaseModel):
    """MCP 输入模型: 查找替换文本"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    search_text: str = Field(
        ...,
        description="Text to search for (max 255 characters, enforced by Word.js API)",
        max_length=255,
    )
    replace_text: str = Field(..., description="Replacement text")
    options: ReplaceOptions | None = Field(
        None,
        description="Replace options (matchCase, matchWholeWord, replaceAll)",
    )


class WordReplaceTextTool(BaseTool):
    """在 Word 文档中查找并替换文本"""

    @property
    def name(self) -> str:
        return "word_replace_text"

    @property
    def description(self) -> str:
        return (
            "Find and replace text in a Word document (equivalent to Ctrl+H). "
            "Searches for the specified text and replaces it with the replacement text. "
            "Supports options for case sensitivity, whole word matching, and replacing all occurrences. "
            "Search text is limited to 255 characters by the Word.js API."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordReplaceTextInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "replace:text"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordReplaceTextInput
