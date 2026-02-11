"""word_get_selection MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class WordGetSelectionInput(BaseModel):
    """MCP 输入模型: 获取选区信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")


class WordGetSelectionTool(BaseTool):
    """获取 Word 文档中的选区位置信息"""

    @property
    def name(self) -> str:
        return "word_get_selection"

    @property
    def description(self) -> str:
        return (
            "Get the current selection information in a Word document. "
            "Returns lightweight position data: selection type, start/end offsets, and selected text. "
            "Use this when you need to know where the cursor is or what range is selected."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetSelectionInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:selection"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetSelectionInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回选区位置摘要 | Get tool: return selection summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        sel_type = obs.data.get("type", "Unknown")
        start = obs.data.get("start", "?")
        end = obs.data.get("end", "?")
        text = obs.data.get("text", "")
        content = f"Selection: type={sel_type}, start={start}, end={end}"
        if text:
            content += f", text={text!r}"
        return {"success": True, "content": content, "data": obs.data}
