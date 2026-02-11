"""word_resolve_comment MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class WordResolveCommentInput(BaseModel):
    """MCP 输入模型: 解决/取消解决批注"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    comment_id: str = Field(..., description="ID of the comment to resolve or unresolve")
    resolved: bool = Field(default=True, description="True to resolve, False to unresolve the comment")


class WordResolveCommentTool(BaseTool):
    """解决或取消解决 Word 文档中的批注"""

    @property
    def name(self) -> str:
        return "word_resolve_comment"

    @property
    def description(self) -> str:
        return (
            "Resolve or unresolve a comment in a Word document. "
            "By default, marks the comment as resolved. Set resolved=false to unresolve. "
            "The comment ID can be obtained from word_get_comments."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordResolveCommentInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "resolve:comment"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordResolveCommentInput
