"""ppt_update_image MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import ImageUpdateOptions, SlideImageData


class PptUpdateImageInput(BaseModel):
    """MCP 输入模型: 替换图片内容"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str = Field(..., description="Image element ID to update")
    image: SlideImageData = Field(..., description="New image data (base64 encoded)")
    options: ImageUpdateOptions | None = Field(None, description="Update options (keepDimensions, width, height)")


class PptUpdateImageTool(BaseTool):
    """替换幻灯片中的图片内容"""

    @property
    def name(self) -> str:
        return "ppt_update_image"

    @property
    def description(self) -> str:
        return (
            "Replace the image content of an existing image element on a PowerPoint slide. "
            "By default keeps the original dimensions. Set keepDimensions=false to resize. "
            "Get the elementId from ppt_get_slide_elements first."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateImageInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:image"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateImageInput
