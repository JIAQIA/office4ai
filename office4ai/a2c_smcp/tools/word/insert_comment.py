"""word_insert_comment MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import InsertCommentTarget


class WordInsertCommentInput(BaseModel):
    """MCP 输入模型: 插入批注"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    text: str = Field(..., description="Comment text content")
    target: InsertCommentTarget | None = Field(
        None,
        description="Target location for the comment (defaults to current selection). "
        "Use type='searchText' with searchText to attach comment to specific text.",
    )


class WordInsertCommentTool(BaseTool):
    """在 Word 文档中插入批注"""

    @property
    def name(self) -> str:
        return "word_insert_comment"

    @property
    def description(self) -> str:
        return (
            "Insert a comment (annotation) into a Word document. "
            "By default, the comment is attached to the current selection. "
            "Optionally specify a target to attach the comment to specific text found by search."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertCommentInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:comment"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertCommentInput
