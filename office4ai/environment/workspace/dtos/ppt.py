"""
PowerPoint Socket.IO DTOs

Defines data structures for PowerPoint-specific Socket.IO events.
"""

from typing import ClassVar, Literal, Optional

from pydantic import BaseModel, Field

from .common import BaseRequest

# ============================================================================
# Content Retrieval DTOs
# ============================================================================


class PptGetCurrentSlideElementsRequest(BaseRequest):
    """
    Request to get current slide elements.
    """

    event_name: ClassVar[str] = "ppt:get:currentSlideElements"

    pass


class PptGetSlideElementsRequest(BaseRequest):
    """
    Request to get elements from a specific slide.
    """

    event_name: ClassVar[str] = "ppt:get:slideElements"

    slideIndex: int = Field(..., description="Slide index (0-based)", ge=0)
    options: Optional["SlideElementsOptions"] = Field(default=None, description="Elements retrieval options")


class PptGetSlideScreenshotRequest(BaseRequest):
    """
    Request to get slide screenshot.
    """

    event_name: ClassVar[str] = "ppt:get:slideScreenshot"

    slideIndex: int = Field(..., description="Slide index (0-based)", ge=0)
    options: Optional["ScreenshotOptions"] = Field(default=None, description="Screenshot options")


class SlideElementsOptions(BaseModel):
    """
    Options for slide elements retrieval.
    """

    includeText: bool = Field(default=True, description="Include text elements")
    includeImages: bool = Field(default=True, description="Include image elements")
    includeShapes: bool = Field(default=True, description="Include shape elements")
    includeTables: bool = Field(default=True, description="Include table elements")
    includeCharts: bool = Field(default=True, description="Include chart elements")


class ScreenshotOptions(BaseModel):
    """
    Screenshot generation options.
    """

    format: Literal["png", "jpeg"] = Field(default="png", description="Image format")
    quality: int = Field(default=90, description="Image quality (1-100)", ge=1, le=100)


# ============================================================================
# Content Insertion DTOs
# ============================================================================


class PptInsertTextRequest(BaseRequest):
    """
    Request to insert text into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:text"

    text: str = Field(..., description="Text to insert")
    options: Optional["TextInsertOptions"] = Field(default=None, description="Insertion options")


class PptInsertImageRequest(BaseRequest):
    """
    Request to insert image into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:image"

    image: "SlideImageData" = Field(..., description="Image data")
    options: Optional["ElementInsertOptions"] = Field(default=None, description="Insertion options")


class PptInsertTableRequest(BaseRequest):
    """
    Request to insert table into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:table"

    options: "SlideTableInsertOptions" = Field(..., description="Table insertion options")


class PptInsertShapeRequest(BaseRequest):
    """
    Request to insert shape into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:shape"

    shapeType: Literal[
        "Rectangle",
        "RoundedRectangle",
        "Circle",
        "Oval",
        "Triangle",
        "Diamond",
        "Pentagon",
        "Hexagon",
        "Line",
        "Arrow",
        "Star",
        "TextBox",
    ] = Field(..., description="Shape type")
    options: Optional["ShapeInsertOptions"] = Field(default=None, description="Insertion options")


class TextInsertOptions(BaseModel):
    """
    Text insertion options.
    """

    slideIndex: int | None = Field(default=None, description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")
    fontSize: int | None = Field(default=None, description="Font size")
    fontName: str | None = Field(default=None, description="Font name")
    color: str | None = Field(default=None, description="Font color (hex)")
    fillColor: str | None = Field(default=None, description="Fill color (hex)")


class SlideImageData(BaseModel):
    """
    Image data for slide insertion.
    """

    base64: str = Field(..., description="Base64 encoded image")


class ElementInsertOptions(BaseModel):
    """
    Common element insertion options.
    """

    slideIndex: int | None = Field(default=None, description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")


class SlideTableInsertOptions(BaseModel):
    """
    Table insertion options for slides.
    """

    rows: int = Field(..., description="Number of rows", ge=1)
    columns: int = Field(..., description="Number of columns", ge=1)
    slideIndex: int | None = Field(default=None, description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    data: list[list[str]] | None = Field(default=None, description="Table data")


class ShapeInsertOptions(BaseModel):
    """
    Shape insertion options.
    """

    slideIndex: int | None = Field(default=None, description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")
    fillColor: str | None = Field(default=None, description="Fill color (hex)")
    borderColor: str | None = Field(default=None, description="Border color (hex)")
    borderWidth: float | None = Field(default=None, description="Border width (points)")
    text: str | None = Field(default=None, description="Shape text")


# ============================================================================
# Slide Management DTOs
# ============================================================================


class PptDeleteSlideRequest(BaseRequest):
    """
    Request to delete a slide.
    """

    event_name: ClassVar[str] = "ppt:delete:slide"

    slideIndex: int = Field(..., description="Slide index (0-based)", ge=0)


class PptMoveSlideRequest(BaseRequest):
    """
    Request to move a slide to a new position.
    """

    event_name: ClassVar[str] = "ppt:move:slide"

    fromIndex: int = Field(..., description="Current slide index (0-based)", ge=0)
    toIndex: int = Field(..., description="Target slide index (0-based)", ge=0)


# ============================================================================
# Update Operation DTOs
# ============================================================================


class PptUpdateTextBoxRequest(BaseRequest):
    """
    Request to update a text box in PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:update:textBox"

    elementId: str = Field(..., description="Element ID to update")
    updates: "TextBoxUpdates" = Field(..., description="Updates to apply")


class TextBoxUpdates(BaseModel):
    """
    Text box update fields.
    """

    text: str | None = Field(default=None, description="New text content")
    fontSize: int | None = Field(default=None, description="Font size")
    fontName: str | None = Field(default=None, description="Font name")
    color: str | None = Field(default=None, description="Font color (hex)")
    fillColor: str | None = Field(default=None, description="Fill color (hex)")
    bold: bool | None = Field(default=None, description="Bold text")
    italic: bool | None = Field(default=None, description="Italic text")


# Resolve forward references
SlideElementsOptions.model_rebuild()
ScreenshotOptions.model_rebuild()
TextInsertOptions.model_rebuild()
SlideImageData.model_rebuild()
ElementInsertOptions.model_rebuild()
SlideTableInsertOptions.model_rebuild()
ShapeInsertOptions.model_rebuild()
TextBoxUpdates.model_rebuild()
