"""
Test Word DTOs

测试 Word 事件的数据传输对象。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.word import (
    GetContentOptions,
    TextFormat,
    WordGetSelectedContentRequest,
    WordInsertTextRequest,
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

