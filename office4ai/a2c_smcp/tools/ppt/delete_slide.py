"""ppt_delete_slide MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class PptDeleteSlideInput(BaseModel):
    """MCP 输入模型: 删除幻灯片"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    slideIndex: int = Field(..., description="Slide index to delete (0-based)", ge=0)


class PptDeleteSlideTool(BaseTool):
    """删除幻灯片"""

    @property
    def name(self) -> str:
        return "ppt_delete_slide"

    @property
    def description(self) -> str:
        return "Delete a slide from a PowerPoint presentation by its index (0-based)."

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptDeleteSlideInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "delete:slide"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptDeleteSlideInput
