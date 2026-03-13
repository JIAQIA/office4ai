"""ppt_insert_image MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import ElementInsertOptions, SlideImageData


class PptInsertImageInput(BaseModel):
    """MCP 输入模型: 插入图片"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    image: SlideImageData = Field(..., description="Image data (base64 encoded)")
    options: ElementInsertOptions | None = Field(None, description="Insertion options (position, size)")


class PptInsertImageTool(BaseTool):
    """在幻灯片上插入图片"""

    @property
    def name(self) -> str:
        return "ppt_insert_image"

    @property
    def description(self) -> str:
        return (
            "Insert an image on a PowerPoint slide. "
            "The image must be provided as a base64-encoded string. "
            "Supports specifying slide index, position, and dimensions."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptInsertImageInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "insert:image"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptInsertImageInput
