"""ppt_get_slide_elements MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.ppt import SlideElementsOptions


class PptGetSlideElementsInput(BaseModel):
    """MCP 输入模型: 获取指定幻灯片元素"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    slideIndex: int = Field(..., description="Slide index (0-based)", ge=0)
    options: SlideElementsOptions | None = Field(None, description="Element filter options")


class PptGetSlideElementsTool(BaseTool):
    """获取指定幻灯片上的所有元素"""

    @property
    def name(self) -> str:
        return "ppt_get_slide_elements"

    @property
    def description(self) -> str:
        return (
            "Get all elements on a specific slide of a PowerPoint presentation. "
            "Supports filtering by element type (text, images, shapes, tables, charts). "
            "Returns element details including type, position, size, text content, and z-order."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetSlideElementsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:slideElements"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetSlideElementsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回元素摘要"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        elements = obs.data.get("elements", [])
        slide_index = obs.data.get("slideIndex", "?")
        content = f"Slide {slide_index}: {len(elements)} element(s)"
        return {"success": True, "content": content, "data": obs.data}
