"""ppt_add_slide MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import AddSlideOptions


class PptAddSlideInput(BaseModel):
    """MCP 输入模型: 添加幻灯片"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    options: AddSlideOptions | None = Field(None, description="Slide options (insertIndex, layout)")


class PptAddSlideTool(BaseTool):
    """添加新幻灯片"""

    @property
    def name(self) -> str:
        return "ppt_add_slide"

    @property
    def description(self) -> str:
        return (
            "Add a new slide to a PowerPoint presentation. "
            "Optionally specify the insert position and layout template name. "
            "Use ppt_get_slide_layouts to discover available layouts first."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptAddSlideInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "add:slide"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptAddSlideInput
