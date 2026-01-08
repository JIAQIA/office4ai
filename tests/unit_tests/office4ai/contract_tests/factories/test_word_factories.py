"""
Test WordDataFactory

测试 Word 数据工厂的各种方法。
"""

from __future__ import annotations

import pytest

from tests.contract_tests.factories.word_factories import WordDataFactory


class TestWordDataFactory:
    """Test WordDataFactory class"""

    @pytest.fixture
    def factory(self) -> WordDataFactory:
        """Create WordDataFactory instance"""
        return WordDataFactory()

    def test_selected_content_response_defaults(self, factory: WordDataFactory) -> None:
        """Test selected_content_response with default values"""
        response = factory.selected_content_response()

        assert "text" in response
        assert "metadata" in response
        assert "elements" in response

        # Verify metadata
        assert response["metadata"]["characterCount"] == len(response["text"])
        assert response["metadata"]["paragraphCount"] == 1
        assert response["metadata"]["tableCount"] == 0
        assert response["metadata"]["imageCount"] == 0

        # Verify elements
        assert len(response["elements"]) == 1
        assert response["elements"][0]["type"] == "paragraph"

    def test_selected_content_response_custom_text(self, factory: WordDataFactory) -> None:
        """Test selected_content_response with custom text"""
        custom_text = "Custom test content"
        response = factory.selected_content_response(text=custom_text)

        assert response["text"] == custom_text
        assert response["metadata"]["characterCount"] == len(custom_text)

    def test_selected_content_response_with_images(self, factory: WordDataFactory) -> None:
        """Test selected_content_response with images"""
        response = factory.selected_content_response(include_images=True)

        assert response["metadata"]["imageCount"] == 1

        # Verify image element exists
        image_elements = [e for e in response["elements"] if e["type"] == "image"]
        assert len(image_elements) == 1
        assert "src" in image_elements[0]
        assert "width" in image_elements[0]
        assert "height" in image_elements[0]

    def test_selected_content_response_with_tables(self, factory: WordDataFactory) -> None:
        """Test selected_content_response with tables"""
        response = factory.selected_content_response(include_tables=True)

        assert response["metadata"]["tableCount"] == 1

        # Verify table element exists
        table_elements = [e for e in response["elements"] if e["type"] == "table"]
        assert len(table_elements) == 1
        assert "rows" in table_elements[0]
        assert "columns" in table_elements[0]
        assert "data" in table_elements[0]

    def test_selected_content_response_without_elements(self, factory: WordDataFactory) -> None:
        """Test selected_content_response without elements"""
        response = factory.selected_content_response(include_elements=False)

        assert response["elements"] == []

    def test_selected_content_request_defaults(self, factory: WordDataFactory) -> None:
        """Test selected_content_request with default values"""
        request = factory.selected_content_request()

        assert "requestId" in request
        assert "documentUri" in request
        assert "options" in request

        # Verify default options
        assert request["options"]["includeText"] is True
        assert request["options"]["includeImages"] is False
        assert request["options"]["includeTables"] is False

    def test_selected_content_request_custom(self, factory: WordDataFactory) -> None:
        """Test selected_content_request with custom values"""
        custom_options = {
            "includeText": True,
            "includeImages": True,
            "includeTables": True,
        }

        request = factory.selected_content_request(
            request_id="custom_req_123",
            document_uri="file:///custom.docx",
            options=custom_options,
        )

        assert request["requestId"] == "custom_req_123"
        assert request["documentUri"] == "file:///custom.docx"
        assert request["options"] == custom_options

    def test_insert_text_response_success(self, factory: WordDataFactory) -> None:
        """Test insert_text_response for success case"""
        response = factory.insert_text_response(success=True)

        assert "insertedLength" in response
        assert "position" in response
        assert response["position"]["start"] == 0
        assert response["position"]["end"] == 11

    def test_insert_text_response_failure(self, factory: WordDataFactory) -> None:
        """Test insert_text_response for failure case"""
        response = factory.insert_text_response(success=False)

        assert response == {}

    def test_insert_text_response_custom_values(self, factory: WordDataFactory) -> None:
        """Test insert_text_response with custom values"""
        response = factory.insert_text_response(
            success=True,
            inserted_length=42,
            position={"start": 10, "end": 52},
        )

        assert response["insertedLength"] == 42
        assert response["position"]["start"] == 10
        assert response["position"]["end"] == 52

    def test_replace_selection_response_success(self, factory: WordDataFactory) -> None:
        """Test replace_selection_response for success case"""
        response = factory.replace_selection_response(success=True)

        assert "replacedLength" in response
        assert response["replacedLength"] == 11

    def test_replace_selection_response_failure(self, factory: WordDataFactory) -> None:
        """Test replace_selection_response for failure case"""
        response = factory.replace_selection_response(success=False)

        assert response == {}

    def test_replace_selection_response_custom_length(self, factory: WordDataFactory) -> None:
        """Test replace_selection_response with custom length"""
        response = factory.replace_selection_response(
            success=True,
            replaced_length=100,
        )

        assert response["replacedLength"] == 100

    def test_error_response_defaults(self, factory: WordDataFactory) -> None:
        """Test error_response with default values"""
        error = factory.error_response()

        assert error["success"] is False
        assert error["error"]["code"] == "3002"
        assert error["error"]["message"] == "Selection is empty"

    def test_error_response_custom(self, factory: WordDataFactory) -> None:
        """Test error_response with custom values"""
        error = factory.error_response(
            code="4000",
            message="Validation error",
        )

        assert error["success"] is False
        assert error["error"]["code"] == "4000"
        assert error["error"]["message"] == "Validation error"

    def test_factory_reusability(self, factory: WordDataFactory) -> None:
        """Test that factory can be reused to generate multiple responses"""
        # Generate multiple responses
        response1 = factory.selected_content_response(text="First")
        response2 = factory.selected_content_response(text="Second")
        response3 = factory.selected_content_response(text="Third")

        # Verify they are independent
        assert response1["text"] == "First"
        assert response2["text"] == "Second"
        assert response3["text"] == "Third"

        # Verify metadata is calculated correctly for each
        assert response1["metadata"]["characterCount"] == 5
        assert response2["metadata"]["characterCount"] == 6
        assert response3["metadata"]["characterCount"] == 5
