"""word_insert_image MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import ImageData, InsertLocation


class WordInsertImageInput(BaseModel):
    """MCP 输入模型: 插入图片"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    image: ImageData = Field(..., description="Image data (base64 encoded, with optional width/height/altText)")
    location: InsertLocation | None = Field(None, description="Insertion location (Cursor, Start, End, etc.)")
    wrap_type: Literal["Inline", "Square", "Tight", "Behind", "InFront"] | None = Field(
        default="Inline",
        description="Text wrapping type around the image",
    )


class WordInsertImageTool(BaseTool):
    """在 Word 文档中插入图片"""

    @property
    def name(self) -> str:
        return "word_insert_image"

    @property
    def description(self) -> str:
        return (
            "Insert an image into a Word document. "
            "The image must be provided as a base64-encoded string. "
            "Supports specifying dimensions, alt text, insertion location, and text wrapping type."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertImageInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:image"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertImageInput
