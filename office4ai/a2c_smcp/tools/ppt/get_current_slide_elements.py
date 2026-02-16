"""ppt_get_current_slide_elements MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class PptGetCurrentSlideElementsInput(BaseModel):
    """MCP 输入模型: 获取当前幻灯片元素"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")


class PptGetCurrentSlideElementsTool(BaseTool):
    """获取当前幻灯片上的所有元素"""

    @property
    def name(self) -> str:
        return "ppt_get_current_slide_elements"

    @property
    def description(self) -> str:
        return (
            "Get all elements on the current slide of a PowerPoint presentation. "
            "Returns element details including type, position, size, text content, and z-order. "
            "Use this to understand the current slide layout before making modifications."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetCurrentSlideElementsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:currentSlideElements"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetCurrentSlideElementsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回元素摘要"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        elements = obs.data.get("elements", [])
        slide_index = obs.data.get("slideIndex", "?")
        content = f"Slide {slide_index}: {len(elements)} element(s)"
        return {"success": True, "content": content, "data": obs.data}
