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


class PptGetSlideInfoRequest(BaseRequest):
    """
    Request to get presentation/slide basic info.
    """

    event_name: ClassVar[str] = "ppt:get:slideInfo"

    slideIndex: int | None = Field(default=None, description="Slide index (0-based), optional", ge=0)


class SlideLayoutsOptions(BaseModel):
    """
    Options for slide layouts retrieval.
    """

    includePlaceholders: bool = Field(default=True, description="Include placeholder details")


class PptGetSlideLayoutsRequest(BaseRequest):
    """
    Request to get available slide layout templates.
    """

    event_name: ClassVar[str] = "ppt:get:slideLayouts"

    options: Optional["SlideLayoutsOptions"] = Field(default=None, description="Layouts retrieval options")


class ImageUpdateOptions(BaseModel):
    """
    Options for image update.
    """

    keepDimensions: bool = Field(default=True, description="Keep original dimensions")
    width: float | None = Field(default=None, description="New width (points)")
    height: float | None = Field(default=None, description="New height (points)")


class PptUpdateImageRequest(BaseRequest):
    """
    Request to replace image content in PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:update:image"

    elementId: str = Field(..., description="Image element ID to update")
    image: "SlideImageData" = Field(..., description="New image data")
    options: Optional["ImageUpdateOptions"] = Field(default=None, description="Update options")


class TableCellUpdate(BaseModel):
    """
    Single table cell update.
    """

    rowIndex: int = Field(..., description="Row index (0-based)", ge=0)
    columnIndex: int = Field(..., description="Column index (0-based)", ge=0)
    text: str = Field(..., description="New text content")


class PptUpdateTableCellRequest(BaseRequest):
    """
    Request to update specific table cells.
    """

    event_name: ClassVar[str] = "ppt:update:tableCell"

    elementId: str = Field(..., description="Table element ID")
    cells: list["TableCellUpdate"] = Field(..., description="Cells to update", min_length=1)


class RowUpdate(BaseModel):
    """
    Row-level batch update.
    """

    rowIndex: int = Field(..., description="Row index (0-based)", ge=0)
    values: list[str] = Field(..., description="Values for each column in the row")


class ColumnUpdate(BaseModel):
    """
    Column-level batch update.
    """

    columnIndex: int = Field(..., description="Column index (0-based)", ge=0)
    values: list[str] = Field(..., description="Values for each row in the column")


class PptUpdateTableRowColumnRequest(BaseRequest):
    """
    Request to batch update table by row or column.
    """

    event_name: ClassVar[str] = "ppt:update:tableRowColumn"

    elementId: str = Field(..., description="Table element ID")
    rows: list["RowUpdate"] | None = Field(default=None, description="Row updates")
    columns: list["ColumnUpdate"] | None = Field(default=None, description="Column updates")


class CellFormat(BaseModel):
    """
    Format for a specific table cell.
    """

    rowIndex: int = Field(..., description="Row index (0-based)", ge=0)
    columnIndex: int = Field(..., description="Column index (0-based)", ge=0)
    backgroundColor: str | None = Field(default=None, description="Background color (hex)")
    fontSize: int | None = Field(default=None, description="Font size")
    fontColor: str | None = Field(default=None, description="Font color (hex)")
    bold: bool | None = Field(default=None, description="Bold text")
    italic: bool | None = Field(default=None, description="Italic text")
    horizontalAlignment: Literal["Left", "Center", "Right"] | None = Field(
        default=None, description="Horizontal alignment"
    )
    verticalAlignment: Literal["Top", "Middle", "Bottom"] | None = Field(default=None, description="Vertical alignment")


class RowFormat(BaseModel):
    """
    Format for a table row.
    """

    rowIndex: int = Field(..., description="Row index (0-based)", ge=0)
    height: float | None = Field(default=None, description="Row height (points)")
    backgroundColor: str | None = Field(default=None, description="Background color (hex)")
    fontSize: int | None = Field(default=None, description="Font size")


class ColumnFormat(BaseModel):
    """
    Format for a table column.
    """

    columnIndex: int = Field(..., description="Column index (0-based)", ge=0)
    width: float | None = Field(default=None, description="Column width (points)")
    backgroundColor: str | None = Field(default=None, description="Background color (hex)")
    fontSize: int | None = Field(default=None, description="Font size")


class PptUpdateTableFormatRequest(BaseRequest):
    """
    Request to update table formatting.
    """

    event_name: ClassVar[str] = "ppt:update:tableFormat"

    elementId: str = Field(..., description="Table element ID")
    cellFormats: list["CellFormat"] | None = Field(default=None, description="Cell-level formats")
    rowFormats: list["RowFormat"] | None = Field(default=None, description="Row-level formats")
    columnFormats: list["ColumnFormat"] | None = Field(default=None, description="Column-level formats")


class ElementUpdates(BaseModel):
    """
    Geometric property updates for an element.
    """

    left: float | None = Field(default=None, description="New X position (points)")
    top: float | None = Field(default=None, description="New Y position (points)")
    width: float | None = Field(default=None, description="New width (points)")
    height: float | None = Field(default=None, description="New height (points)")
    rotation: float | None = Field(default=None, description="New rotation angle (degrees, 0-360)")


class PptUpdateElementRequest(BaseRequest):
    """
    Request to update element position/size/rotation.
    """

    event_name: ClassVar[str] = "ppt:update:element"

    elementId: str = Field(..., description="Element ID to update")
    slideIndex: int | None = Field(default=None, description="Slide index (0-based)", ge=0)
    updates: "ElementUpdates" = Field(..., description="Geometric updates")


class PptDeleteElementRequest(BaseRequest):
    """
    Request to delete element(s) from a slide.
    """

    event_name: ClassVar[str] = "ppt:delete:element"

    elementId: str | None = Field(default=None, description="Single element ID to delete")
    elementIds: list[str] | None = Field(default=None, description="Batch element IDs to delete")
    slideIndex: int | None = Field(default=None, description="Slide index (0-based)", ge=0)


class PptReorderElementRequest(BaseRequest):
    """
    Request to adjust element z-order.
    """

    event_name: ClassVar[str] = "ppt:reorder:element"

    elementId: str = Field(..., description="Element ID to reorder")
    slideIndex: int | None = Field(default=None, description="Slide index (0-based)", ge=0)
    action: Literal["bringToFront", "sendToBack", "bringForward", "sendBackward"] = Field(
        ..., description="Reorder action"
    )


class AddSlideOptions(BaseModel):
    """
    Options for adding a new slide.
    """

    insertIndex: int | None = Field(default=None, description="Insert position index (0-based)", ge=0)
    layout: str | None = Field(default=None, description="Layout name (e.g. 'Title Slide', 'Blank')")


class PptAddSlideRequest(BaseRequest):
    """
    Request to add a new slide.
    """

    event_name: ClassVar[str] = "ppt:add:slide"

    options: Optional["AddSlideOptions"] = Field(default=None, description="Slide options")


class PptGotoSlideRequest(BaseRequest):
    """
    Request to jump to a specific slide.
    """

    event_name: ClassVar[str] = "ppt:goto:slide"

    slideIndex: int = Field(..., description="Target slide index (0-based)", ge=0)


# Resolve forward references
SlideElementsOptions.model_rebuild()
ScreenshotOptions.model_rebuild()
TextInsertOptions.model_rebuild()
SlideImageData.model_rebuild()
ElementInsertOptions.model_rebuild()
SlideTableInsertOptions.model_rebuild()
ShapeInsertOptions.model_rebuild()
TextBoxUpdates.model_rebuild()
SlideLayoutsOptions.model_rebuild()
ImageUpdateOptions.model_rebuild()
TableCellUpdate.model_rebuild()
RowUpdate.model_rebuild()
ColumnUpdate.model_rebuild()
CellFormat.model_rebuild()
RowFormat.model_rebuild()
ColumnFormat.model_rebuild()
ElementUpdates.model_rebuild()
AddSlideOptions.model_rebuild()
