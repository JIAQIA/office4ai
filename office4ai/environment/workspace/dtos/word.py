"""
Word Socket.IO DTOs

Defines data structures for Word-specific Socket.IO events.
"""

from typing import Any, Literal, Optional

from pydantic import BaseModel, Field

from .common import BaseRequest

# ============================================================================
# Content Retrieval DTOs
# ============================================================================


class WordGetSelectedContentRequest(BaseRequest):
    """
    Request to get selected content from Word document.
    """

    options: Optional["GetContentOptions"] = Field(default=None, description="Content retrieval options")


class WordGetVisibleContentRequest(BaseRequest):
    """
    Request to get visible content from Word document.
    """

    options: Optional["GetContentOptions"] = Field(default=None, description="Content retrieval options")


class WordGetDocumentStructureRequest(BaseRequest):
    """
    Request to get Word document structure.
    """

    pass


class WordGetDocumentStatsRequest(BaseRequest):
    """
    Request to get Word document statistics.
    """

    pass


class GetContentOptions(BaseModel):
    """
    Options for content retrieval operations.
    """

    includeText: bool = Field(default=True, description="Include text content")
    includeImages: bool = Field(default=True, description="Include images")
    includeTables: bool = Field(default=True, description="Include tables")
    includeContentControls: bool = Field(default=False, description="Include content controls")
    detailedMetadata: bool = Field(default=False, description="Include detailed metadata")
    maxTextLength: int | None = Field(default=None, description="Maximum text length")


# ============================================================================
# Text Operation DTOs
# ============================================================================


class WordInsertTextRequest(BaseRequest):
    """
    Request to insert text into Word document.
    """

    text: str = Field(..., description="Text to insert")
    location: Literal["Cursor", "Start", "End"] = Field(default="Cursor", description="Insertion location")
    format: Optional["TextFormat"] = Field(default=None, description="Text formatting")


class WordReplaceSelectionRequest(BaseRequest):
    """
    Request to replace selected content in Word document.
    """

    content: "ReplaceContent" = Field(..., description="Replacement content")


class WordReplaceTextRequest(BaseRequest):
    """
    Request to find and replace text in Word document.
    """

    searchText: str = Field(..., description="Text to search for")
    replaceText: str = Field(..., description="Replacement text")
    options: Optional["ReplaceOptions"] = Field(default=None, description="Replace options")


class WordAppendTextRequest(BaseRequest):
    """
    Request to append text to Word document.
    """

    text: str = Field(..., description="Text to append")
    location: Literal["Start", "End"] = Field(default="End", description="Append location")
    format: Optional["TextFormat"] = Field(default=None, description="Text formatting")


class TextFormat(BaseModel):
    """
    Text formatting options.
    """

    bold: bool | None = Field(default=None, description="Bold text")
    italic: bool | None = Field(default=None, description="Italic text")
    fontSize: int | None = Field(default=None, description="Font size")
    fontName: str | None = Field(default=None, description="Font name")
    color: str | None = Field(default=None, description="Font color (hex)")
    underline: bool | None = Field(default=None, description="Underline text")


class ReplaceContent(BaseModel):
    """
    Content for replacement operation.
    """

    text: str | None = Field(default=None, description="Replacement text")
    images: list[dict[str, Any]] | None = Field(default=None, description="Replacement images")
    format: Optional["TextFormat"] = Field(default=None, description="Text formatting")


class ReplaceOptions(BaseModel):
    """
    Options for find and replace operation.
    """

    matchCase: bool = Field(default=False, description="Match case")
    matchWholeWord: bool = Field(default=False, description="Match whole word")
    replaceAll: bool = Field(default=False, description="Replace all occurrences")


# ============================================================================
# Multimedia Operation DTOs
# ============================================================================


class WordInsertImageRequest(BaseRequest):
    """
    Request to insert image into Word document.
    """

    image: "ImageData" = Field(..., description="Image data")
    location: Optional["InsertLocation"] = Field(default=None, description="Insertion location")
    wrapType: Literal["Inline", "Square", "Tight", "Behind", "InFront"] | None = Field(
        default="Inline", description="Text wrapping type"
    )


class WordInsertTableRequest(BaseRequest):
    """
    Request to insert table into Word document.
    """

    options: "TableInsertOptions" = Field(..., description="Table insertion options")


class WordInsertEquationRequest(BaseRequest):
    """
    Request to insert equation into Word document.
    """

    latex: str = Field(..., description="LaTeX equation string")
    options: Optional["EquationOptions"] = Field(default=None, description="Equation options")


class ImageData(BaseModel):
    """
    Image data for insertion.
    """

    base64: str = Field(..., description="Base64 encoded image")
    width: int | None = Field(default=None, description="Image width")
    height: int | None = Field(default=None, description="Image height")
    altText: str | None = Field(default=None, description="Alternative text")


class InsertLocation(BaseModel):
    """
    Insertion location details.
    """

    type: Literal["Cursor", "Start", "End", "BeforeBookmark", "AfterBookmark"] = Field(
        ...,
        description="Location type",
    )
    bookmarkName: str | None = Field(default=None, description="Bookmark name if applicable")


class TableInsertOptions(BaseModel):
    """
    Table insertion options.
    """

    rows: int = Field(..., description="Number of rows", ge=1)
    columns: int = Field(..., description="Number of columns", ge=1)
    data: list[list[str]] | None = Field(default=None, description="Table data")
    style: str | None = Field(default=None, description="Table style name")


class EquationOptions(BaseModel):
    """
    Equation insertion options.
    """

    inline: bool = Field(default=True, description="Inline equation")


# ============================================================================
# Advanced Feature DTOs
# ============================================================================


class WordInsertTOCRequest(BaseRequest):
    """
    Request to insert table of contents into Word document.
    """

    options: Optional["TOCOptions"] = Field(default=None, description="TOC options")


class WordExportContentRequest(BaseRequest):
    """
    Request to export content from Word document.
    """

    format: Literal["text", "html", "markdown"] = Field(..., description="Export format")
    options: Optional["ExportOptions"] = Field(default=None, description="Export options")


class TOCOptions(BaseModel):
    """
    Table of contents options.
    """

    maxLevel: int = Field(default=3, description="Maximum heading level", ge=1, le=9)
    styles: list[str] | None = Field(default=None, description="Heading styles to include")


class ExportOptions(BaseModel):
    """
    Content export options.
    """

    includeImages: bool = Field(default=False, description="Include images in export")
    includeTables: bool = Field(default=True, description="Include tables in export")


# Resolve forward references
GetContentOptions.model_rebuild()
TextFormat.model_rebuild()
ReplaceContent.model_rebuild()
ReplaceOptions.model_rebuild()
ImageData.model_rebuild()
InsertLocation.model_rebuild()
TableInsertOptions.model_rebuild()
EquationOptions.model_rebuild()
TOCOptions.model_rebuild()
ExportOptions.model_rebuild()
