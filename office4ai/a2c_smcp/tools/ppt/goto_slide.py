"""ppt_goto_slide MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class PptGotoSlideInput(BaseModel):
    """MCP 输入模型: 跳转到幻灯片"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    slideIndex: int = Field(..., description="Target slide index (0-based)", ge=0)


class PptGotoSlideTool(BaseTool):
    """跳转到指定幻灯片"""

    @property
    def name(self) -> str:
        return "ppt_goto_slide"

    @property
    def description(self) -> str:
        return "Jump to a specific slide in a PowerPoint presentation, making it the current displayed slide."

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGotoSlideInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "goto:slide"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGotoSlideInput
