"""ppt_get_slide_info MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class PptGetSlideInfoInput(BaseModel):
    """MCP 输入模型: 获取演示文稿/幻灯片基本信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    slideIndex: int | None = Field(
        None, description="Slide index (0-based). If specified, returns detailed slide info", ge=0
    )


class PptGetSlideInfoTool(BaseTool):
    """获取演示文稿/幻灯片基本信息"""

    @property
    def name(self) -> str:
        return "ppt_get_slide_info"

    @property
    def description(self) -> str:
        return (
            "Get presentation metadata: total slide count, dimensions, aspect ratio, and current slide index. "
            "Optionally specify a slideIndex to get detailed info about that slide including layout, elements, and background. "
            "Use this before layout calculations to know the slide dimensions and element distribution."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetSlideInfoInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:slideInfo"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetSlideInfoInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回演示文稿信息摘要"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        slide_count = obs.data.get("slideCount", 0)
        dimensions = obs.data.get("dimensions", {})
        width = dimensions.get("width", "?")
        height = dimensions.get("height", "?")
        aspect = dimensions.get("aspectRatio", "?")
        content = f"{slide_count} slides, {width}x{height} pt ({aspect})"
        return {"success": True, "content": content, "data": obs.data}
