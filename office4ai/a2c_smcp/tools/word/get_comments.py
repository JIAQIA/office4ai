"""word_get_comments MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.word import GetCommentsOptions


class WordGetCommentsInput(BaseModel):
    """MCP 输入模型: 获取批注"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    options: GetCommentsOptions | None = Field(
        None,
        description="Comment retrieval options (includeResolved, includeReplies, includeAssociatedText, detailedMetadata)",
    )


class WordGetCommentsTool(BaseTool):
    """获取 Word 文档中的批注"""

    @property
    def name(self) -> str:
        return "word_get_comments"

    @property
    def description(self) -> str:
        return (
            "Get comments (annotations) from a Word document. "
            "Returns a list of comments with their content, author, and resolution status. "
            "Supports filtering resolved comments and including reply threads."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetCommentsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:comments"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetCommentsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回批注摘要 | Get tool: return comments summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        comments = obs.data.get("comments", [])
        if not comments:
            content = "No comments found"
        else:
            summaries = []
            for c in comments:
                author = c.get("authorName", "Unknown")
                text = c.get("content", "")
                resolved = c.get("resolved", False)
                status = " [resolved]" if resolved else ""
                summaries.append(f"- {author}{status}: {text}")
            content = f"{len(comments)} comment(s):\n" + "\n".join(summaries)
        return {"success": True, "content": content, "data": obs.data}
