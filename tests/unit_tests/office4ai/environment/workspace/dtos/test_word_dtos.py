"""
Test Word DTOs

测试 Word 事件的数据传输对象。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.word import (
    AnyContentElement,
    ContentMetadata,
    GetContentOptions,
    GetStylesOptions,
    ReplaceContent,
    StyleInfo,
    StylesResult,
    TextFormat,
    WordGetSelectedContentRequest,
    WordGetSelectedContentResponse,
    WordGetStylesRequest,
    WordGetVisibleContentRequest,
    WordGetVisibleContentResponse,
    WordInsertTextRequest,
    WordReplaceSelectionRequest,
    WordReplaceSelectionResponse,
)


class TestWordGetSelectedContentRequest:
    """Test WordGetSelectedContentRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetSelectedContentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetContentOptions(
            includeText=True,
            includeImages=False,
            includeTables=True,
        )

        request = WordGetSelectedContentRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.include_images is False
        assert request.options.include_tables is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordGetSelectedContentRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            options={"includeText": True, "detailedMetadata": True},
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.detailed_metadata is True

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetSelectedContentRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetSelectedContentRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetSelectedContentRequest.event_name == "word:get:selectedContent"


class TestGetContentOptions:
    """Test GetContentOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = GetContentOptions()

        assert options.include_text is True
        assert options.include_images is True
        assert options.include_tables is True
        assert options.include_content_controls is False
        assert options.detailed_metadata is False
        assert options.max_text_length is None

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = GetContentOptions(
            includeText=False,
            includeImages=False,
            includeTables=False,
            includeContentControls=True,
            detailedMetadata=True,
            maxTextLength=1000,
        )

        assert options.include_text is False
        assert options.include_images is False
        assert options.include_tables is False
        assert options.include_content_controls is True
        assert options.detailed_metadata is True
        assert options.max_text_length == 1000

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = GetContentOptions(
            **{
                "includeText": True,
                "includeImages": False,
                "maxTextLength": 500,
            }
        )

        assert options.include_text is True
        assert options.include_images is False
        assert options.max_text_length == 500

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = GetContentOptions(
            includeText=True,
            includeImages=True,
            detailedMetadata=True,
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = options.model_dump(by_alias=True)

        assert data["includeText"] is True
        assert data["includeImages"] is True
        assert data["detailedMetadata"] is True
        # Default values should be included
        assert "includeTables" in data

    def test_invalid_max_text_length(self) -> None:
        """Test that negative max_text_length is rejected"""
        # Pydantic doesn't have a constraint on max_text_length, so this should work
        # In a real scenario, you might want to add a validation constraint
        options = GetContentOptions(maxTextLength=-100)
        assert options.max_text_length == -100


class TestWordInsertTextRequest:
    """Test WordInsertTextRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordInsertTextRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            text="Hello World",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.text == "Hello World"
        assert request.location == "Cursor"  # Default value
        assert request.format is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_location(self) -> None:
        """Test creating valid request with different locations"""
        # Test "Start" location
        request_start = WordInsertTextRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            text="Text at start",
            location="Start",
        )
        assert request_start.location == "Start"

        # Test "End" location
        request_end = WordInsertTextRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            text="Text at end",
            location="End",
        )
        assert request_end.location == "End"

        # Test "Cursor" location (default)
        request_cursor = WordInsertTextRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            text="Text at cursor",
            location="Cursor",
        )
        assert request_cursor.location == "Cursor"

    def test_valid_request_with_format(self) -> None:
        """Test creating valid request with text format"""
        format_options = TextFormat(
            bold=True,
            italic=False,
            fontSize=14,
            fontName="Arial",
            color="#FF0000",
        )

        request = WordInsertTextRequest(
            requestId="req_005",
            documentUri="file:///test.docx",
            text="Formatted text",
            location="Cursor",
            format=format_options,
        )

        assert request.text == "Formatted text"
        assert request.format is not None
        assert request.format.bold is True
        assert request.format.italic is False
        assert request.format.font_size == 14
        assert request.format.font_name == "Arial"
        assert request.format.color == "#FF0000"

    def test_request_with_dict_format(self) -> None:
        """Test creating request with format as dict"""
        request = WordInsertTextRequest(
            requestId="req_006",
            documentUri="file:///test.docx",
            text="Text with format",
            format={"bold": True, "italic": True, "fontSize": 12},
        )

        assert request.format is not None
        assert request.format.bold is True
        assert request.format.italic is True
        assert request.format.font_size == 12

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                documentUri="file:///test.docx",
                text="Hello",
            )

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                requestId="req_001",
                text="Hello",
            )

        assert "documentUri" in str(exc_info.value)

        # Missing text
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "text" in str(exc_info.value)

    def test_invalid_location(self) -> None:
        """Test validation fails with invalid location"""
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
                text="Hello",
                location="InvalidLocation",  # type: ignore
            )

        assert "location" in str(exc_info.value).lower()

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordInsertTextRequest.event_name == "word:insert:text"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordInsertTextRequest(
            requestId="req_007",
            documentUri="file:///test.docx",
            text="Hello",
            location="Start",
            format={"bold": True},
        )

        payload = request.to_payload()

        assert payload["requestId"] == "req_007"
        assert payload["documentUri"] == "file:///test.docx"
        assert payload["text"] == "Hello"
        assert payload["location"] == "Start"
        assert payload["format"]["bold"] is True
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        """Test build class method for creating requests"""
        request = WordInsertTextRequest.build(
            document_uri="file:///test.docx",
            text="Auto-generated request",
            location="End",
        )

        assert request.request_id is not None
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///test.docx"
        assert request.text == "Auto-generated request"
        assert request.location == "End"


class TestTextFormat:
    """Test TextFormat DTO"""

    def test_default_format(self) -> None:
        """Test creating format with default values"""
        format_options = TextFormat()

        assert format_options.bold is None
        assert format_options.italic is None
        assert format_options.font_size is None
        assert format_options.font_name is None
        assert format_options.color is None
        assert format_options.underline is None

    def test_custom_format(self) -> None:
        """Test creating format with custom values"""
        format_options = TextFormat(
            bold=True,
            italic=True,
            fontSize=16,
            fontName="Times New Roman",
            color="#0000FF",
            underline=True,
        )

        assert format_options.bold is True
        assert format_options.italic is True
        assert format_options.font_size == 16
        assert format_options.font_name == "Times New Roman"
        assert format_options.color == "#0000FF"
        assert format_options.underline is True

    def test_partial_format(self) -> None:
        """Test creating format with partial fields"""
        format_options = TextFormat(bold=True, fontSize=14)

        assert format_options.bold is True
        assert format_options.italic is None
        assert format_options.font_size == 14
        assert format_options.font_name is None

    def test_format_serialization(self) -> None:
        """Test format can be serialized to dict with correct aliases"""
        format_options = TextFormat(
            bold=True,
            italic=False,
            fontSize=12,
            fontName="Arial",
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = format_options.model_dump(by_alias=True, exclude_none=True)

        assert data["bold"] is True
        assert data["italic"] is False
        assert data["fontSize"] == 12
        assert data["fontName"] == "Arial"
        # None values should be excluded with exclude_none=True
        assert "color" not in data
        assert "underline" not in data

    def test_format_from_dict(self) -> None:
        """Test creating format from dict with aliases"""
        format_options = TextFormat(
            **{
                "bold": True,
                "italic": True,
                "fontSize": 14,
                "color": "#FF0000",
            }
        )

        assert format_options.bold is True
        assert format_options.italic is True
        assert format_options.font_size == 14
        assert format_options.color == "#FF0000"


class TestWordGetStylesRequest:
    """Test WordGetStylesRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetStylesRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetStylesOptions(
            includeBuiltIn=True,
            includeCustom=True,
            includeUnused=False,
            detailedInfo=True,
        )

        request = WordGetStylesRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_built_in is True
        assert request.options.include_custom is True
        assert request.options.include_unused is False
        assert request.options.detailed_info is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordGetStylesRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            options={"includeBuiltIn": False, "includeCustom": True, "detailedInfo": False},
        )

        assert request.options is not None
        assert request.options.include_built_in is False
        assert request.options.include_custom is True
        assert request.options.detailed_info is False

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetStylesRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetStylesRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetStylesRequest.event_name == "word:get:styles"


class TestGetStylesOptions:
    """Test GetStylesOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = GetStylesOptions()

        assert options.include_built_in is True
        assert options.include_custom is True
        assert options.include_unused is False
        assert options.detailed_info is False

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = GetStylesOptions(
            includeBuiltIn=False,
            includeCustom=True,
            includeUnused=True,
            detailedInfo=True,
        )

        assert options.include_built_in is False
        assert options.include_custom is True
        assert options.include_unused is True
        assert options.detailed_info is True

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = GetStylesOptions(
            **{
                "includeBuiltIn": False,
                "includeCustom": True,
                "detailedInfo": True,
            }
        )

        assert options.include_built_in is False
        assert options.include_custom is True
        assert options.detailed_info is True
        # include_unused should use default value
        assert options.include_unused is False

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = GetStylesOptions(
            includeBuiltIn=True,
            includeCustom=True,
            detailedInfo=True,
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = options.model_dump(by_alias=True)

        assert data["includeBuiltIn"] is True
        assert data["includeCustom"] is True
        assert data["detailedInfo"] is True
        # Default values should be included
        assert "includeUnused" in data


class TestStyleInfo:
    """Test StyleInfo DTO"""

    def test_valid_style_info(self) -> None:
        """Test creating valid style info"""
        style = StyleInfo(
            name="标题一",
            type="Paragraph",
            builtIn=True,
            inUse=True,
        )

        assert style.name == "标题一"
        assert style.type == "Paragraph"
        assert style.built_in is True
        assert style.in_use is True
        assert style.description is None

    def test_style_info_with_description(self) -> None:
        """Test creating style info with description"""
        style = StyleInfo(
            name="Normal",
            type="Paragraph",
            builtIn=True,
            inUse=True,
            description="Normal paragraph style",
        )

        assert style.name == "Normal"
        assert style.description == "Normal paragraph style"

    def test_invalid_style_type(self) -> None:
        """Test validation fails with invalid style type"""
        with pytest.raises(ValidationError) as exc_info:
            StyleInfo(
                name="Test Style",
                type="InvalidType",  # type: ignore
                builtIn=True,
                inUse=True,
            )

        assert "type" in str(exc_info.value).lower()

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing name
        with pytest.raises(ValidationError) as exc_info:
            StyleInfo(
                type="Paragraph",  # type: ignore
                builtIn=True,
                inUse=True,
            )

        assert "name" in str(exc_info.value)

        # Missing type
        with pytest.raises(ValidationError) as exc_info:
            StyleInfo(
                name="Test",
                builtIn=True,
                inUse=True,
            )

        assert "type" in str(exc_info.value)

    def test_all_style_types(self) -> None:
        """Test all valid style types"""
        types = ["Paragraph", "Character", "Table", "List"]

        for style_type in types:
            style = StyleInfo(
                name=f"Test {style_type}",
                type=style_type,  # type: ignore
                builtIn=True,
                inUse=True,
            )
            assert style.type == style_type

    def test_style_serialization(self) -> None:
        """Test style can be serialized to dict with correct aliases"""
        style = StyleInfo(
            name="标题一",
            type="Paragraph",
            builtIn=True,
            inUse=True,
            description="Heading style",
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = style.model_dump(by_alias=True)

        assert data["name"] == "标题一"
        assert data["type"] == "Paragraph"
        assert data["builtIn"] is True
        assert data["inUse"] is True
        assert data["description"] == "Heading style"


class TestStylesResult:
    """Test StylesResult DTO"""

    def test_valid_styles_result(self) -> None:
        """Test creating valid styles result"""
        styles = [
            StyleInfo(name="标题一", type="Paragraph", builtIn=True, inUse=True),
            StyleInfo(name="正文", type="Paragraph", builtIn=True, inUse=True),
            StyleInfo(name="强调", type="Character", builtIn=True, inUse=False),
        ]

        result = StylesResult(styles=styles)

        assert len(result.styles) == 3
        assert result.styles[0].name == "标题一"
        assert result.styles[1].name == "正文"
        assert result.styles[2].name == "强调"

    def test_empty_styles_result(self) -> None:
        """Test creating empty styles result"""
        result = StylesResult(styles=[])

        assert len(result.styles) == 0

    def test_styles_result_serialization(self) -> None:
        """Test styles result can be serialized to dict with correct aliases"""
        styles = [
            StyleInfo(
                name="Normal",
                type="Paragraph",
                builtIn=True,
                inUse=True,
                description="Normal style",
            ),
        ]

        result = StylesResult(styles=styles)
        data = result.model_dump(by_alias=True)

        assert "styles" in data
        assert len(data["styles"]) == 1
        assert data["styles"][0]["name"] == "Normal"
        assert data["styles"][0]["type"] == "Paragraph"
        assert data["styles"][0]["description"] == "Normal style"

    def test_styles_result_from_dict(self) -> None:
        """Test creating styles result from dict"""
        styles_data = [
            {
                "name": "标题一",
                "type": "Paragraph",
                "builtIn": True,
                "inUse": True,
            },
            {
                "name": "Custom Style",
                "type": "Character",
                "builtIn": False,
                "inUse": True,
            },
        ]

        result = StylesResult(**{"styles": styles_data})

        assert len(result.styles) == 2
        assert result.styles[0].name == "标题一"
        assert result.styles[1].name == "Custom Style"
        assert result.styles[1].built_in is False


class TestWordReplaceSelectionRequest:
    """Test WordReplaceSelectionRequest DTO"""

    def test_valid_request_with_text_content(self) -> None:
        """Test creating valid request with text content"""
        content = ReplaceContent(text="Replaced text")

        request = WordReplaceSelectionRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            content=content,
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.content.text == "Replaced text"
        assert request.content.images is None
        assert request.content.format is None

    def test_valid_request_with_images_content(self) -> None:
        """Test creating valid request with images content"""
        images = [{"base64": "iVBORw0KGgo...", "width": 100, "height": 100}]
        content = ReplaceContent(images=images)

        request = WordReplaceSelectionRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            content=content,
        )

        assert request.content.images == images
        assert request.content.text is None

    def test_valid_request_with_text_and_format(self) -> None:
        """Test creating valid request with text and format"""
        text_format = TextFormat(bold=True, italic=True, fontSize=14)
        content = ReplaceContent(text="Formatted text", format=text_format)

        request = WordReplaceSelectionRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            content=content,
        )

        assert request.content.text == "Formatted text"
        assert request.content.format is not None
        assert request.content.format.bold is True
        assert request.content.format.italic is True
        assert request.content.format.font_size == 14

    def test_request_with_dict_content(self) -> None:
        """Test creating request with content as dict"""
        request = WordReplaceSelectionRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            content={"text": "Test", "format": {"bold": True}},
        )

        assert request.content.text == "Test"
        assert request.content.format is not None
        assert request.content.format.bold is True

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceSelectionRequest(
                documentUri="file:///test.docx",
                content={"text": "Test"},
            )

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceSelectionRequest(
                requestId="req_001",
                content={"text": "Test"},
            )

        assert "documentUri" in str(exc_info.value)

        # Missing content
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceSelectionRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "content" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordReplaceSelectionRequest.event_name == "word:replace:selection"

    def test_request_serialization(self) -> None:
        """Test request can be serialized to dict with correct aliases"""
        content = ReplaceContent(text="Test content")
        request = WordReplaceSelectionRequest(
            requestId="req_005",
            documentUri="file:///test.docx",
            content=content,
        )

        data = request.model_dump(by_alias=True)

        assert data["requestId"] == "req_005"
        assert data["documentUri"] == "file:///test.docx"
        assert data["content"]["text"] == "Test content"


class TestWordReplaceSelectionResponse:
    """Test WordReplaceSelectionResponse DTO"""

    def test_valid_response_success(self) -> None:
        """Test creating valid success response"""
        response = WordReplaceSelectionResponse(
            replaced=True,
            characterCount=100,
        )

        assert response.replaced is True
        assert response.character_count == 100

    def test_valid_response_failure(self) -> None:
        """Test creating valid failure response"""
        response = WordReplaceSelectionResponse(
            replaced=False,
            characterCount=0,
        )

        assert response.replaced is False
        assert response.character_count == 0

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        response = WordReplaceSelectionResponse(
            replaced=True,
            characterCount=50,
        )

        data = response.model_dump(by_alias=True)

        assert data["replaced"] is True
        assert data["characterCount"] == 50

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {"replaced": True, "characterCount": 200}

        response = WordReplaceSelectionResponse(**data)

        assert response.replaced is True
        assert response.character_count == 200


class TestReplaceContent:
    """Test ReplaceContent DTO"""

    def test_content_with_text_only(self) -> None:
        """Test creating content with text only"""
        content = ReplaceContent(text="Some text")

        assert content.text == "Some text"
        assert content.images is None
        assert content.format is None

    def test_content_with_images_only(self) -> None:
        """Test creating content with images only"""
        images = [
            {"base64": "abc123", "width": 200, "height": 150, "altText": "Image 1"},
            {"base64": "def456", "width": 300, "height": 200},
        ]
        content = ReplaceContent(images=images)

        assert content.images == images
        assert content.text is None
        assert content.format is None

    def test_content_with_text_and_format(self) -> None:
        """Test creating content with text and format"""
        format_obj = TextFormat(
            bold=True,
            italic=False,
            fontSize=16,
            fontName="Arial",
            color="#FF0000",
        )
        content = ReplaceContent(text="Bold red text", format=format_obj)

        assert content.text == "Bold red text"
        assert content.format is not None
        assert content.format.bold is True
        assert content.format.italic is False
        assert content.format.font_size == 16
        assert content.format.font_name == "Arial"
        assert content.format.color == "#FF0000"

    def test_content_serialization(self) -> None:
        """Test content can be serialized to dict with correct aliases"""
        format_obj = TextFormat(bold=True, fontSize=14)
        content = ReplaceContent(text="Test", format=format_obj)

        data = content.model_dump(by_alias=True)

        assert data["text"] == "Test"
        assert data["format"]["bold"] is True
        assert data["format"]["fontSize"] == 14


class TestWordGetSelectedContentResponse:
    """Test WordGetSelectedContentResponse DTO"""

    def test_valid_response(self) -> None:
        """Test creating valid response"""
        metadata = ContentMetadata(isEmpty=False, characterCount=100)
        response = WordGetSelectedContentResponse(
            text="Selected text content",
            elements=[],
            metadata=metadata,
        )

        assert response.text == "Selected text content"
        assert response.elements == []
        assert response.metadata is not None
        assert response.metadata.is_empty is False
        assert response.metadata.character_count == 100

    def test_response_without_metadata(self) -> None:
        """Test creating response without metadata"""
        response = WordGetSelectedContentResponse(
            text="Text without metadata",
            elements=[],
        )

        assert response.text == "Text without metadata"
        assert response.metadata is None

    def test_response_with_elements(self) -> None:
        """Test creating response with elements"""
        elements = [
            AnyContentElement(type="text", content={"text": "Paragraph 1"}),
            AnyContentElement(type="image", content={"base64": "abc123", "width": 100}),
        ]
        response = WordGetSelectedContentResponse(
            text="Content with elements",
            elements=elements,
        )

        assert len(response.elements) == 2
        assert response.elements[0].type == "text"
        assert response.elements[1].type == "image"

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        metadata = ContentMetadata(isEmpty=False, characterCount=50)
        response = WordGetSelectedContentResponse(
            text="Test content",
            elements=[],
            metadata=metadata,
        )

        data = response.model_dump(by_alias=True)

        assert data["text"] == "Test content"
        assert data["metadata"]["isEmpty"] is False
        assert data["metadata"]["characterCount"] == 50


class TestWordGetVisibleContentRequest:
    """Test WordGetVisibleContentRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetVisibleContentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetContentOptions(
            includeText=True,
            includeImages=False,
            includeTables=True,
        )

        request = WordGetVisibleContentRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.include_images is False
        assert request.options.include_tables is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordGetVisibleContentRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            options={"includeText": True, "maxTextLength": 1000},
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.max_text_length == 1000

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetVisibleContentRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetVisibleContentRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetVisibleContentRequest.event_name == "word:get:visibleContent"


class TestWordGetVisibleContentResponse:
    """Test WordGetVisibleContentResponse DTO"""

    def test_valid_response(self) -> None:
        """Test creating valid response"""
        metadata = ContentMetadata(isEmpty=False, characterCount=200)
        response = WordGetVisibleContentResponse(
            text="Visible content text",
            elements=[],
            metadata=metadata,
        )

        assert response.text == "Visible content text"
        assert response.elements == []
        assert response.metadata is not None
        assert response.metadata.is_empty is False
        assert response.metadata.character_count == 200

    def test_empty_content_response(self) -> None:
        """Test creating empty content response"""
        metadata = ContentMetadata(isEmpty=True, characterCount=0)
        response = WordGetVisibleContentResponse(
            text="",
            elements=[],
            metadata=metadata,
        )

        assert response.text == ""
        assert response.metadata.is_empty is True
        assert response.metadata.character_count == 0

    def test_response_with_complex_elements(self) -> None:
        """Test creating response with various element types"""
        elements = [
            AnyContentElement(type="text", content={"text": "Heading"}),
            AnyContentElement(type="image", content={"base64": "iVBORw0KGgo...", "width": 200}),
            AnyContentElement(type="table", content={"rows": 3, "columns": 2}),
            AnyContentElement(type="other", content={"type": "equation"}),
        ]
        response = WordGetVisibleContentResponse(
            text="Complex content",
            elements=elements,
        )

        assert len(response.elements) == 4
        element_types = [elem.type for elem in response.elements]
        assert "text" in element_types
        assert "image" in element_types
        assert "table" in element_types
        assert "other" in element_types

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        metadata = ContentMetadata(isEmpty=False, characterCount=150)
        response = WordGetVisibleContentResponse(
            text="Serialized content",
            elements=[AnyContentElement(type="text", content={"text": "Test"})],
            metadata=metadata,
        )

        data = response.model_dump(by_alias=True)

        assert data["text"] == "Serialized content"
        assert len(data["elements"]) == 1
        assert data["elements"][0]["type"] == "text"
        assert data["metadata"]["isEmpty"] is False
        assert data["metadata"]["characterCount"] == 150


class TestContentMetadata:
    """Test ContentMetadata DTO"""

    def test_valid_metadata(self) -> None:
        """Test creating valid metadata"""
        metadata = ContentMetadata(isEmpty=False, characterCount=100)

        assert metadata.is_empty is False
        assert metadata.character_count == 100

    def test_empty_metadata(self) -> None:
        """Test creating empty content metadata"""
        metadata = ContentMetadata(isEmpty=True, characterCount=0)

        assert metadata.is_empty is True
        assert metadata.character_count == 0

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing isEmpty
        with pytest.raises(ValidationError) as exc_info:
            ContentMetadata(characterCount=100)

        assert "isEmpty" in str(exc_info.value)

        # Missing characterCount
        with pytest.raises(ValidationError) as exc_info:
            ContentMetadata(isEmpty=False)

        assert "characterCount" in str(exc_info.value)

    def test_metadata_serialization(self) -> None:
        """Test metadata can be serialized to dict with correct aliases"""
        metadata = ContentMetadata(isEmpty=False, characterCount=500)

        data = metadata.model_dump(by_alias=True)

        assert data["isEmpty"] is False
        assert data["characterCount"] == 500


class TestAnyContentElement:
    """Test AnyContentElement DTO"""

    def test_text_element(self) -> None:
        """Test creating text element"""
        element = AnyContentElement(type="text", content={"text": "Paragraph text"})

        assert element.type == "text"
        assert element.content["text"] == "Paragraph text"

    def test_image_element(self) -> None:
        """Test creating image element"""
        element = AnyContentElement(
            type="image",
            content={"base64": "iVBORw0KGgo...", "width": 300, "height": 200},
        )

        assert element.type == "image"
        assert element.content["width"] == 300
        assert element.content["height"] == 200

    def test_table_element(self) -> None:
        """Test creating table element"""
        element = AnyContentElement(
            type="table",
            content={"rows": 3, "columns": 4, "data": [["A", "B"], ["C", "D"]]},
        )

        assert element.type == "table"
        assert element.content["rows"] == 3
        assert element.content["columns"] == 4

    def test_other_element(self) -> None:
        """Test creating other type element"""
        element = AnyContentElement(type="other", content={"type": "equation", "latex": "x^2"})

        assert element.type == "other"
        assert element.content["latex"] == "x^2"

    def test_invalid_element_type(self) -> None:
        """Test validation fails with invalid element type"""
        with pytest.raises(ValidationError) as exc_info:
            AnyContentElement(
                type="invalid_type",  # type: ignore
                content={},
            )

        assert "type" in str(exc_info.value).lower()

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing type
        with pytest.raises(ValidationError) as exc_info:
            AnyContentElement(content={})  # type: ignore

        assert "type" in str(exc_info.value)

        # Missing content
        with pytest.raises(ValidationError) as exc_info:
            AnyContentElement(type="text", content=None)  # type: ignore

        assert "content" in str(exc_info.value)

    def test_element_serialization(self) -> None:
        """Test element can be serialized to dict with correct aliases"""
        element = AnyContentElement(type="text", content={"text": "Test text"})

        data = element.model_dump(by_alias=True)

        assert data["type"] == "text"
        assert data["content"]["text"] == "Test text"

    def test_all_valid_element_types(self) -> None:
        """Test all valid element types"""
        types = ["text", "image", "table", "other"]

        for element_type in types:
            element = AnyContentElement(type=element_type, content={"data": "test"})
            assert element.type == element_type


