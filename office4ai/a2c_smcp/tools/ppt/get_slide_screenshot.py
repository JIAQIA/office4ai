"""ppt_get_slide_screenshot MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.ppt import ScreenshotOptions


class PptGetSlideScreenshotInput(BaseModel):
    """MCP 输入模型: 获取幻灯片截图"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    slideIndex: int = Field(..., description="Slide index (0-based)", ge=0)
    options: ScreenshotOptions | None = Field(None, description="Screenshot options (format, quality)")


class PptGetSlideScreenshotTool(BaseTool):
    """获取幻灯片截图"""

    @property
    def name(self) -> str:
        return "ppt_get_slide_screenshot"

    @property
    def description(self) -> str:
        return (
            "Get a screenshot of a specific slide as a Base64-encoded image. "
            "Supports PNG and JPEG formats with configurable quality. "
            "Use this to visually inspect a slide's current appearance."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetSlideScreenshotInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:slideScreenshot"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetSlideScreenshotInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回截图信息"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        fmt = obs.data.get("format", "unknown")
        base64_data = obs.data.get("base64", "")
        content = f"Screenshot ({fmt}): {len(base64_data)} chars base64"
        return {"success": True, "content": content, "data": obs.data}
