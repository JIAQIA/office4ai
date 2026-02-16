"""ppt_delete_element MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class PptDeleteElementInput(BaseModel):
    """MCP 输入模型: 删除元素"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | None = Field(None, description="Single element ID to delete")
    elementIds: list[str] | None = Field(None, description="Batch element IDs to delete")
    slideIndex: int | None = Field(None, description="Slide index (0-based), default: current slide", ge=0)


class PptDeleteElementTool(BaseTool):
    """删除幻灯片上的元素"""

    @property
    def name(self) -> str:
        return "ppt_delete_element"

    @property
    def description(self) -> str:
        return (
            "Delete one or more elements from a PowerPoint slide. "
            "Provide elementId for single deletion or elementIds for batch deletion. "
            "If both are provided, elementIds takes priority."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptDeleteElementInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "delete:element"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptDeleteElementInput
