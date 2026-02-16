"""ppt_get_slide_layouts MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.ppt import SlideLayoutsOptions


class PptGetSlideLayoutsInput(BaseModel):
    """MCP 输入模型: 获取可用幻灯片版式列表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    options: SlideLayoutsOptions | None = Field(None, description="Layouts retrieval options")


class PptGetSlideLayoutsTool(BaseTool):
    """获取可用幻灯片版式列表"""

    @property
    def name(self) -> str:
        return "ppt_get_slide_layouts"

    @property
    def description(self) -> str:
        return (
            "Get all available slide layout templates in the presentation. "
            "Returns layout names, types, placeholder counts, and whether they are custom. "
            "Use this before adding slides to know which layouts are available."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetSlideLayoutsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:slideLayouts"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetSlideLayoutsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回版式列表摘要"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        layouts = obs.data.get("layouts", [])
        if not layouts:
            content = "No layouts found"
        else:
            names = [layout.get("name", "?") for layout in layouts]
            content = f"{len(layouts)} layout(s): {', '.join(names)}"
        return {"success": True, "content": content, "data": obs.data}
