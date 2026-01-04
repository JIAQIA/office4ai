"""
Base Namespace

Base functionality for all Socket.IO namespaces.
"""

import logging
from typing import Any

from socketio import AsyncNamespace  # type: ignore[import-untyped]

from office4ai.environment.workspace.socketio.services.connection_manager import ClientInfo, connection_manager

logger = logging.getLogger(__name__)


class BaseNamespace(AsyncNamespace):
    """
    Base namespace class with common functionality.

    All namespaces (/word, /ppt, /excel) inherit from this.
    """

    def __init__(self, namespace: str) -> None:
        """
        Initialize base namespace.

        Args:
            namespace: Namespace name (/word, /ppt, /excel)
            ns_name: Optional namespace name (defaults to namespace)
        """
        super().__init__(namespace)
        self.namespace_name: str = namespace
        logger.info(f"BaseNamespace initialized: {self.namespace_name}")

    async def on_connect(self, sid: str, data: Any) -> None:
        """
        Handle client connection.

        Args:
            sid: Session ID
            data: Handshake data (clientId, documentUri)
        """
        client_id = data.get("clientId") if data else None
        document_uri = data.get("documentUri") if data else None

        if not client_id or not document_uri:
            logger.error(f"Connection failed: missing handshake data from {sid}")
            # Disconnect client
            await self.disconnect(sid)
            return

        # Register client
        try:
            connection_manager.register_client(
                socket_id=sid,
                client_id=client_id,
                document_uri=document_uri,
                namespace=self.namespace_name,
            )

            # Send confirmation
            await self.emit(
                "connection:established",
                {
                    "sessionId": sid,
                    "status": "ready",
                    "serverTime": int(connection_manager.get_connection_count() * 1000),
                },
                to=sid,
            )

            logger.info(f"Client connected: {client_id} ({sid}) for {document_uri} on {self.namespace_name}")

        except Exception as e:
            logger.error(f"Error registering client: {e}", exc_info=True)
            await self.disconnect(sid)

    async def on_disconnect(self, sid: str) -> None:
        """
        Handle client disconnection.

        Args:
            sid: Session ID
        """
        client_info = connection_manager.unregister_client(sid)

        if client_info:
            logger.info(f"Client disconnected: {client_info.client_id} ({sid}) from {client_info.document_uri}")
        else:
            logger.warning(f"Unknown client disconnected: {sid}")

    async def on_connection_status(self, sid: str, data: Any) -> None:
        """
        Handle connection status updates from client.

        Args:
            sid: Session ID
            data: Status data
        """
        logger.debug(f"Connection status from {sid}: {data}")
        # Can be used for health checks
        # No response needed (fire-and-forget)

    def get_client_info(self, sid: str) -> ClientInfo | None:
        """
        Get client information.

        Args:
            sid: Session ID

        Returns:
            ClientInfo if found, None otherwise
        """
        return connection_manager.get_client_info(sid)
