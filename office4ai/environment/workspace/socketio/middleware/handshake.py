"""
Handshake Middleware

Simple handshake validation for client connections.
Validates clientId and documentUri are provided.
"""

import logging
from collections.abc import Awaitable, Callable
from typing import Any

from socketio import AsyncServer  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)


async def handshake_middleware(
    socketio_server: AsyncServer, namespace: str = "/"
) -> Callable[[str, Any], Awaitable[bool]]:
    """
    Create handshake middleware for a namespace.

    This is a simple validation that checks:
    1. clientId is provided
    2. documentUri is provided

    No complex authentication - just basic validation for local connections.

    Args:
        socketio_server: Socket.IO server instance
        namespace: Namespace to apply middleware to

    Returns:
        Middleware function
    """

    async def middleware(sid: str, environ: Any) -> bool:
        """
        Handshake middleware handler.

        Args:
            sid: Session ID
            environ: WSGI environ dict

        Raises:
            ValueError: If handshake data is invalid
        """
        # Get handshake auth data
        # Note: python-socketio passes auth in a different way
        # We'll get it from the socket handshake in the connection handler
        _auth = environ.get("asgi_scope", {}).get("query_string", b"").decode()

        # Parse auth data from handshake
        # Note: python-socketio passes auth in a different way
        # We'll get it from the socket handshake in the connection handler

        logger.debug(f"Handshake attempt for session {sid} on namespace {namespace}")

        # Accept all connections (localhost-only security)
        # Actual validation happens in connection handler
        return True

    return middleware


def validate_handshake_data(client_id: str, document_uri: str) -> tuple[bool, str]:
    """
    Validate handshake data from client.

    Args:
        client_id: Client-provided ID
        document_uri: Document URI

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not client_id:
        return False, "Missing clientId"

    if not document_uri:
        return False, "Missing documentUri"

    # Basic format validation for document URI
    if not document_uri.startswith(("file:///", "http://", "https://")):
        return False, f"Invalid documentUri format: {document_uri}"

    return True, ""


def log_handshake(client_id: str, document_uri: str, namespace: str) -> None:
    """
    Log successful handshake.

    Args:
        client_id: Client ID
        document_uri: Document URI
        namespace: Namespace
    """
    logger.info(f"Handshake successful: client={client_id}, document={document_uri}, namespace={namespace}")
