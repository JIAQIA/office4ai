"""ppt_update_text_box MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import TextBoxUpdates


class PptUpdateTextBoxInput(BaseModel):
    """MCP 输入模型: 更新文本框"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str = Field(..., description="Element ID to update")
    updates: TextBoxUpdates = Field(..., description="Updates to apply (text, fontSize, fontName, color, bold, italic)")


class PptUpdateTextBoxTool(BaseTool):
    """更新幻灯片中的文本框内容/样式"""

    @property
    def name(self) -> str:
        return "ppt_update_text_box"

    @property
    def description(self) -> str:
        return (
            "Update a text box, placeholder, or geometric shape's text content and styling on a PowerPoint slide. "
            "Supports changing text, font size, font name, color, fill color, bold, and italic. "
            "Get the elementId from ppt_get_slide_elements or ppt_get_current_slide_elements first."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateTextBoxInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:textBox"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateTextBoxInput
