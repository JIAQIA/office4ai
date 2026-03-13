"""word_select_text MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import SelectTextSearchOptions


class WordSelectTextInput(BaseModel):
    """MCP 输入模型: 搜索并选中文本"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    search_text: str = Field(
        ...,
        description=(
            "Text to search for (max 255 characters). "
            "For invisible characters, use Word notation: "
            "^p for paragraph break (Enter), ^l for line break (Shift+Enter), ^t for tab. "
            "Do NOT use \\n or \\t — they will not match."
        ),
        max_length=255,
    )
    search_options: SelectTextSearchOptions | None = Field(
        None,
        description="Search options (matchCase, matchWholeWord, matchWildcards)",
    )
    selection_mode: Literal["select", "start", "end"] = Field(
        default="select",
        description="Selection mode: select/highlight text, start cursor at beginning, or end cursor at end",
    )
    select_index: int = Field(
        default=1,
        description="Which match to select (1-based, default: 1)",
        ge=1,
    )


class WordSelectTextTool(BaseTool):
    """在 Word 文档中搜索并选中文本"""

    @property
    def name(self) -> str:
        return "word_select_text"

    @property
    def description(self) -> str:
        return (
            "Search for text in a Word document and select/highlight a specific match. "
            "Supports case-sensitive search, whole word matching, and wildcard patterns. "
            "Can select the Nth occurrence and position the cursor at start, end, or select the full match. "
            "Search text is limited to 255 characters by the Word.js API."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordSelectTextInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "select:text"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordSelectTextInput
