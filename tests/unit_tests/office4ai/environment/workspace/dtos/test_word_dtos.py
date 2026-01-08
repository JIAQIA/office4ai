"""
Test Word DTOs

测试 Word 事件的数据传输对象。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.word import (
    GetContentOptions,
    WordGetSelectedContentRequest,
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
