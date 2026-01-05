"""
Test handshake middleware

测试握手中间件验证功能。
"""

from collections.abc import Awaitable, Callable
from typing import Any

import pytest

from office4ai.environment.workspace.socketio.middleware.handshake import (
    validate_handshake_data,
    log_handshake,
    handshake_middleware,
)


class TestHandshakeMiddleware:
    """Test handshake validation functions"""

    def test_validate_valid_handshake(self) -> None:
        """Test validation with valid data"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="file:///tmp/test.docx")

        assert is_valid is True
        assert error == ""

    def test_validate_missing_client_id(self) -> None:
        """Test validation fails without clientId"""
        is_valid, error = validate_handshake_data(client_id="", document_uri="file:///tmp/test.docx")

        assert is_valid is False
        assert "Missing clientId" in error

    def test_validate_missing_document_uri(self) -> None:
        """Test validation fails without documentUri"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="")

        assert is_valid is False
        assert "Missing documentUri" in error

    def test_validate_invalid_uri_format(self) -> None:
        """Test validation fails with invalid URI format"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="invalid-uri")

        assert is_valid is False
        assert "Invalid documentUri format" in error

    def test_validate_file_uri(self) -> None:
        """Test validation accepts file:// URI"""
        is_valid, error = validate_handshake_data(
            client_id="test_client", document_uri="file:///C:/Users/test.docx"
        )

        assert is_valid is True

    def test_validate_http_uri(self) -> None:
        """Test validation accepts http/https URIs"""
        is_valid, error = validate_handshake_data(
            client_id="test_client", document_uri="https://example.com/test.docx"
        )

        assert is_valid is True

    def test_log_handshake(self, caplog: pytest.LogCaptureFixture) -> None:
        """Test handshake logging"""
        import logging

        caplog.set_level(logging.INFO)

        log_handshake("client1", "file:///test.docx", "/word")

        # Verify log message was created
        assert len(caplog.records) > 0
        assert "client1" in caplog.text
        assert "file:///test.docx" in caplog.text

    @pytest.mark.asyncio
    async def test_handshake_middleware(self) -> None:
        """Test handshake middleware function"""
        from socketio import AsyncServer  # type: ignore[import-untyped]

        # Create a mock server
        server = AsyncServer(async_mode="aiohttp", logger=False, engineio_logger=False)

        # Get middleware for /word namespace
        middleware = await handshake_middleware(server, "/word")

        # Test middleware accepts connection
        environ: dict[str, Any] = {
            "asgi_scope": {
                "query_string": b"",
            }
        }

        result = await middleware("test_sid", environ)

        # Middleware should accept all connections (localhost security)
        assert result is True
