"""
Word Socket.IO DTOs

Defines data structures for Word-specific Socket.IO events.
"""

from typing import Any, ClassVar, Literal, Optional

from pydantic import Field

from .common import BaseRequest, SocketIOBaseModel

# ============================================================================
# Content Retrieval DTOs
# ============================================================================


class WordGetStylesRequest(BaseRequest):
    """
    Request to get all available styles from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:get:styles"

    options: Optional["GetStylesOptions"] = Field(
        default=None,
        alias="options",
        description="Style retrieval options",
    )


class WordGetSelectedContentRequest(BaseRequest):
    """
    Request to get selected content from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:get:selectedContent"

    options: Optional["GetContentOptions"] = Field(
        default=None,
        alias="options",
        description="Content retrieval options",
    )


class WordGetSelectedContentResponse(SocketIOBaseModel):
    """
    Response for word:get:selectedContent operation.

    Uses Pydantic aliases for protocol compliance.
    """

    text: str = Field(..., alias="text", description="Selected text content")
    elements: list["AnyContentElement"] = Field(
        default_factory=list,
        alias="elements",
        description="Content elements (text, images, tables)",
    )
    metadata: Optional["ContentMetadata"] = Field(
        default=None,
        alias="metadata",
        description="Content metadata",
    )


class WordGetVisibleContentRequest(BaseRequest):
    """
    Request to get visible content from Word document.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30736386
    """

    event_name: ClassVar[str] = "word:get:visibleContent"

    options: Optional["GetContentOptions"] = Field(
        default=None,
        alias="options",
        description="Content retrieval options",
    )


class WordGetVisibleContentResponse(SocketIOBaseModel):
    """
    Response for word:get:visibleContent operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30736386
    """

    text: str = Field(..., alias="text", description="Visible text content")
    elements: list["AnyContentElement"] = Field(
        default_factory=list,
        alias="elements",
        description="Content elements (text, images, tables)",
    )
    metadata: Optional["ContentMetadata"] = Field(
        default=None,
        alias="metadata",
        description="Content metadata",
    )


class WordGetDocumentStructureRequest(BaseRequest):
    """
    Request to get Word document structure.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153
    """

    event_name: ClassVar[str] = "word:get:documentStructure"


class DocumentStructure(SocketIOBaseModel):
    """
    Document structure information.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153
    """

    paragraph_count: int = Field(
        ...,
        alias="paragraphCount",
        description="Number of paragraphs in the document",
    )
    table_count: int = Field(
        ...,
        alias="tableCount",
        description="Number of tables in the document",
    )
    image_count: int = Field(
        ...,
        alias="imageCount",
        description="Number of images in the document",
    )
    section_count: int = Field(
        ...,
        alias="sectionCount",
        description="Number of sections in the document",
    )


class WordGetDocumentStructureResponse(SocketIOBaseModel):
    """
    Response for word:get:documentStructure operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153
    """

    data: DocumentStructure = Field(
        ...,
        alias="data",
        description="Document structure information",
    )


class WordGetDocumentStatsRequest(BaseRequest):
    """
    Request to get Word document statistics.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938
    """

    event_name: ClassVar[str] = "word:get:documentStats"


class DocumentStats(SocketIOBaseModel):
    """
    Document statistics information.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938
    """

    word_count: int = Field(
        ...,
        alias="wordCount",
        description="Number of words in the document",
    )
    character_count: int = Field(
        ...,
        alias="characterCount",
        description="Number of characters in the document (including spaces and punctuation)",
    )
    paragraph_count: int = Field(
        ...,
        alias="paragraphCount",
        description="Number of paragraphs in the document (including empty paragraphs)",
    )


class WordGetDocumentStatsResponse(SocketIOBaseModel):
    """
    Response for word:get:documentStats operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938
    """

    data: DocumentStats = Field(
        ...,
        alias="data",
        description="Document statistics information",
    )


class GetContentOptions(SocketIOBaseModel):
    """
    Options for content retrieval operations.

    Uses Pydantic aliases for protocol compliance.
    """

    include_text: bool = Field(
        default=True,
        alias="includeText",
        description="Include text content",
    )
    include_images: bool = Field(
        default=True,
        alias="includeImages",
        description="Include images",
    )
    include_tables: bool = Field(
        default=True,
        alias="includeTables",
        description="Include tables",
    )
    include_content_controls: bool = Field(
        default=False,
        alias="includeContentControls",
        description="Include content controls",
    )
    detailed_metadata: bool = Field(
        default=False,
        alias="detailedMetadata",
        description="Include detailed metadata",
    )
    max_text_length: int | None = Field(
        default=None,
        alias="maxTextLength",
        description="Maximum text length",
    )


class ContentMetadata(SocketIOBaseModel):
    """
    Metadata for content retrieval results.

    Uses Pydantic aliases for protocol compliance.
    """

    is_empty: bool = Field(
        ...,
        alias="isEmpty",
        description="Whether the content is empty",
    )
    character_count: int = Field(
        ...,
        alias="characterCount",
        description="Number of characters in the content",
    )


class AnyContentElement(SocketIOBaseModel):
    """
    Generic content element for different content types.

    Uses Pydantic aliases for protocol compliance.
    """

    type: Literal["text", "image", "table", "other"] = Field(
        ...,
        alias="type",
        description="Element type",
    )
    content: dict[str, Any] = Field(
        ...,
        alias="content",
        description="Element content (varies by type)",
    )


# ============================================================================
# Text Operation DTOs
# ============================================================================


class WordInsertTextRequest(BaseRequest):
    """
    Request to insert text into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:text"

    text: str = Field(..., alias="text", description="Text to insert")
    location: Literal["Cursor", "Start", "End"] = Field(
        default="Cursor",
        alias="location",
        description="Insertion location",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting",
    )


class WordReplaceSelectionRequest(BaseRequest):
    """
    Request to replace selected content in Word document.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30605313
    """

    event_name: ClassVar[str] = "word:replace:selection"

    content: "ReplaceContent" = Field(
        ...,
        alias="content",
        description="Replacement content (text or images must be provided)",
    )


class WordReplaceSelectionResponse(SocketIOBaseModel):
    """
    Response for word:replace:selection operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30605313
    """

    replaced: bool = Field(..., alias="replaced", description="Whether content was replaced")
    character_count: int = Field(..., alias="characterCount", description="Number of characters replaced")


class WordReplaceTextRequest(BaseRequest):
    """
    Request to find and replace text in Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:replace:text"

    search_text: str = Field(
        ...,
        alias="searchText",
        description="Text to search for",
    )
    replace_text: str = Field(
        ...,
        alias="replaceText",
        description="Replacement text",
    )
    options: Optional["ReplaceOptions"] = Field(
        default=None,
        alias="options",
        description="Replace options",
    )


class WordAppendTextRequest(BaseRequest):
    """
    Request to append text to Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:append:text"

    text: str = Field(..., alias="text", description="Text to append")
    location: Literal["Start", "End"] = Field(
        default="End",
        alias="location",
        description="Append location",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting",
    )


class TextFormat(SocketIOBaseModel):
    """
    Text formatting options.

    Uses Pydantic aliases for protocol compliance.
    """

    bold: bool | None = Field(default=None, alias="bold", description="Bold text")
    italic: bool | None = Field(
        default=None,
        alias="italic",
        description="Italic text",
    )
    font_size: int | None = Field(
        default=None,
        alias="fontSize",
        description="Font size",
    )
    font_name: str | None = Field(
        default=None,
        alias="fontName",
        description="Font name",
    )
    color: str | None = Field(
        default=None,
        alias="color",
        description="Font color (hex)",
    )
    underline: bool | None = Field(
        default=None,
        alias="underline",
        description="Underline text",
    )


class ReplaceContent(SocketIOBaseModel):
    """
    Content for replacement operation.

    Uses Pydantic aliases for protocol compliance.
    """

    text: str | None = Field(
        default=None,
        alias="text",
        description="Replacement text",
    )
    images: list[dict[str, Any]] | None = Field(
        default=None,
        alias="images",
        description="Replacement images",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting",
    )


class ReplaceOptions(SocketIOBaseModel):
    """
    Options for find and replace operation.

    Uses Pydantic aliases for protocol compliance.
    """

    match_case: bool = Field(
        default=False,
        alias="matchCase",
        description="Match case",
    )
    match_whole_word: bool = Field(
        default=False,
        alias="matchWholeWord",
        description="Match whole word",
    )
    replace_all: bool = Field(
        default=False,
        alias="replaceAll",
        description="Replace all occurrences",
    )


# ============================================================================
# Multimedia Operation DTOs
# ============================================================================


class WordInsertImageRequest(BaseRequest):
    """
    Request to insert image into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:image"

    image: "ImageData" = Field(..., alias="image", description="Image data")
    location: Optional["InsertLocation"] = Field(
        default=None,
        alias="location",
        description="Insertion location",
    )
    wrap_type: Literal["Inline", "Square", "Tight", "Behind", "InFront"] | None = Field(
        default="Inline",
        alias="wrapType",
        description="Text wrapping type",
    )


class WordInsertTableRequest(BaseRequest):
    """
    Request to insert table into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:table"

    options: "TableInsertOptions" = Field(
        ...,
        alias="options",
        description="Table insertion options",
    )


class WordInsertEquationRequest(BaseRequest):
    """
    Request to insert equation into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:equation"

    latex: str = Field(..., alias="latex", description="LaTeX equation string")
    options: Optional["EquationOptions"] = Field(
        default=None,
        alias="options",
        description="Equation options",
    )


class ImageData(SocketIOBaseModel):
    """
    Image data for insertion.

    Uses Pydantic aliases for protocol compliance.
    """

    base64: str = Field(..., alias="base64", description="Base64 encoded image")
    width: int | None = Field(
        default=None,
        alias="width",
        description="Image width",
    )
    height: int | None = Field(
        default=None,
        alias="height",
        description="Image height",
    )
    alt_text: str | None = Field(
        default=None,
        alias="altText",
        description="Alternative text",
    )


class InsertLocation(SocketIOBaseModel):
    """
    Insertion location details.

    Uses Pydantic aliases for protocol compliance.
    """

    type: Literal["Cursor", "Start", "End", "BeforeBookmark", "AfterBookmark"] = Field(
        ...,
        alias="type",
        description="Location type",
    )
    bookmark_name: str | None = Field(
        default=None,
        alias="bookmarkName",
        description="Bookmark name if applicable",
    )


class TableInsertOptions(SocketIOBaseModel):
    """
    Table insertion options.

    Uses Pydantic aliases for protocol compliance.
    """

    rows: int = Field(..., alias="rows", description="Number of rows", ge=1)
    columns: int = Field(..., alias="columns", description="Number of columns", ge=1)
    data: list[list[str]] | None = Field(
        default=None,
        alias="data",
        description="Table data",
    )
    style: str | None = Field(
        default=None,
        alias="style",
        description="Table style name",
    )


class EquationOptions(SocketIOBaseModel):
    """
    Equation insertion options.

    Uses Pydantic aliases for protocol compliance.
    """

    inline: bool = Field(
        default=True,
        alias="inline",
        description="Inline equation",
    )


# ============================================================================
# Advanced Feature DTOs
# ============================================================================


class WordInsertTOCRequest(BaseRequest):
    """
    Request to insert table of contents into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:toc"

    options: Optional["TOCOptions"] = Field(
        default=None,
        alias="options",
        description="TOC options",
    )


class WordExportContentRequest(BaseRequest):
    """
    Request to export content from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:export:content"

    format: Literal["text", "html", "markdown"] = Field(
        ...,
        alias="format",
        description="Export format",
    )
    options: Optional["ExportOptions"] = Field(
        default=None,
        alias="options",
        description="Export options",
    )


class TOCOptions(SocketIOBaseModel):
    """
    Table of contents options.

    Uses Pydantic aliases for protocol compliance.
    """

    max_level: int = Field(
        default=3,
        alias="maxLevel",
        description="Maximum heading level",
        ge=1,
        le=9,
    )
    styles: list[str] | None = Field(
        default=None,
        alias="styles",
        description="Heading styles to include",
    )


class ExportOptions(SocketIOBaseModel):
    """
    Content export options.

    Uses Pydantic aliases for protocol compliance.
    """

    include_images: bool = Field(
        default=False,
        alias="includeImages",
        description="Include images in export",
    )
    include_tables: bool = Field(
        default=True,
        alias="includeTables",
        description="Include tables in export",
    )


# ============================================================================
# Style DTOs
# ============================================================================


class GetStylesOptions(SocketIOBaseModel):
    """
    Options for retrieving styles from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    include_built_in: bool = Field(
        default=True,
        alias="includeBuiltIn",
        description="Include built-in styles",
    )
    include_custom: bool = Field(
        default=True,
        alias="includeCustom",
        description="Include custom styles",
    )
    include_unused: bool = Field(
        default=False,
        alias="includeUnused",
        description="Include unused styles",
    )
    detailed_info: bool = Field(
        default=False,
        alias="detailedInfo",
        description="Include detailed information like description",
    )


class StyleInfo(SocketIOBaseModel):
    """
    Information about a Word style.

    Uses Pydantic aliases for protocol compliance.
    """

    name: str = Field(..., alias="name", description="Style name (localized)")
    type: Literal["Paragraph", "Character", "Table", "List"] = Field(
        ...,
        alias="type",
        description="Style type",
    )
    built_in: bool = Field(..., alias="builtIn", description="Whether it's a built-in style")
    in_use: bool = Field(..., alias="inUse", description="Whether the style is used in document")
    description: str | None = Field(
        default=None,
        alias="description",
        description="Style description (only when detailedInfo=true)",
    )


class StylesResult(SocketIOBaseModel):
    """
    Result containing styles from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    styles: list[StyleInfo] = Field(
        ...,
        alias="styles",
        description="List of styles",
    )


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
GetStylesOptions.model_rebuild()
StyleInfo.model_rebuild()
StylesResult.model_rebuild()
WordReplaceSelectionResponse.model_rebuild()
ContentMetadata.model_rebuild()
AnyContentElement.model_rebuild()
WordGetSelectedContentResponse.model_rebuild()
WordGetVisibleContentResponse.model_rebuild()
DocumentStructure.model_rebuild()
WordGetDocumentStructureResponse.model_rebuild()
DocumentStats.model_rebuild()
WordGetDocumentStatsResponse.model_rebuild()
