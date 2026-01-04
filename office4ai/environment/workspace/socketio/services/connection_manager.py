"""
Connection Manager

Manages client connections and mappings between documentUri and socketId.
This is the core routing component for Workspace → Add-In communication.
"""

import logging
import time
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class ClientInfo:
    """
    Information about a connected client.
    """

    socket_id: str
    client_id: str
    document_uri: str
    namespace: str  # /word, /ppt, /excel
    connected_at: float  # timestamp

    def __hash__(self) -> int:
        return hash(self.socket_id)


class ConnectionManager:
    """
    Manages active connections and document URI mappings.

    Core functionality:
    1. Track connected clients
    2. Map documentUri → socketId for routing
    3. Handle connection/disconnection events
    """

    def __init__(self) -> None:
        # socket_id → ClientInfo
        self._clients: dict[str, ClientInfo] = {}

        # document_uri → set of socket_ids (one document can have multiple connections)
        self._document_to_sockets: dict[str, set[str]] = {}

        # client_id → socket_id (for tracking unique clients)
        self._client_id_to_socket: dict[str, str] = {}

        logger.info("ConnectionManager initialized")

    def register_client(
        self,
        socket_id: str,
        client_id: str,
        document_uri: str,
        namespace: str,
    ) -> ClientInfo:
        """
        Register a new client connection.

        Args:
            socket_id: Socket.IO socket ID
            client_id: Client-provided unique identifier
            document_uri: Document URI being worked on
            namespace: Namespace (/word, /ppt, /excel)

        Returns:
            ClientInfo: Registered client information
        """
        logger.info("Client registered", extra={"client_id": client_id, "document_uri": document_uri})
        client_info = ClientInfo(
            socket_id=socket_id,
            client_id=client_id,
            document_uri=document_uri,
            namespace=namespace,
            connected_at=time.time(),
        )

        # Store client info
        self._clients[socket_id] = client_info
        self._client_id_to_socket[client_id] = socket_id

        # Map document_uri → socket_id
        if document_uri not in self._document_to_sockets:
            self._document_to_sockets[document_uri] = set()
        self._document_to_sockets[document_uri].add(socket_id)

        logger.info(f"Client registered: {client_id} ({socket_id}) for {document_uri} on {namespace}")

        return client_info

    def unregister_client(self, socket_id: str) -> ClientInfo | None:
        """
        Unregister a client connection (cleanup on disconnect).

        Args:
            socket_id: Socket.IO socket ID

        Returns:
            ClientInfo if found, None otherwise
        """
        client_info = self._clients.pop(socket_id, None)
        if not client_info:
            logger.warning(f"Client not found: {socket_id}")
            return None

        # Remove from client_id mapping
        self._client_id_to_socket.pop(client_info.client_id, None)

        # Remove from document_uri mapping
        if client_info.document_uri in self._document_to_sockets:
            self._document_to_sockets[client_info.document_uri].discard(socket_id)
            if not self._document_to_sockets[client_info.document_uri]:
                # No more connections for this document
                del self._document_to_sockets[client_info.document_uri]

        logger.info(f"Client unregistered: {client_info.client_id} ({socket_id}) for {client_info.document_uri}")

        return client_info

    def get_socket_by_document(self, document_uri: str) -> str | None:
        """
        Get socket ID for a document URI.

        Args:
            document_uri: Document URI

        Returns:
            Socket ID if found, None otherwise
        """
        sockets = self._document_to_sockets.get(document_uri)
        if not sockets:
            return None

        # Return the first active socket for this document
        # (typically there should be only one)
        return next(iter(sockets), None)

    def get_client_info(self, socket_id: str) -> ClientInfo | None:
        """
        Get client information by socket ID.

        Args:
            socket_id: Socket.IO socket ID

        Returns:
            ClientInfo if found, None otherwise
        """
        return self._clients.get(socket_id)

    def get_all_clients(self) -> list[ClientInfo]:
        """
        Get all connected clients.

        Returns:
            List of ClientInfo objects
        """
        return list(self._clients.values())

    def get_clients_by_document(self, document_uri: str) -> list[ClientInfo]:
        """
        Get all clients working on a document.

        Args:
            document_uri: Document URI

        Returns:
            List of ClientInfo objects
        """
        socket_ids = self._document_to_sockets.get(document_uri, set())
        return [self._clients[sid] for sid in socket_ids if sid in self._clients]

    def get_clients_by_namespace(self, namespace: str) -> list[ClientInfo]:
        """
        Get all clients in a namespace.

        Args:
            namespace: Namespace (/word, /ppt, /excel)

        Returns:
            List of ClientInfo objects
        """
        return [client for client in self._clients.values() if client.namespace == namespace]

    def is_document_active(self, document_uri: str) -> bool:
        """
        Check if a document has active connections.

        Args:
            document_uri: Document URI

        Returns:
            True if document has active connections, False otherwise
        """
        return document_uri in self._document_to_sockets and bool(self._document_to_sockets[document_uri])

    def get_connection_count(self) -> int:
        """
        Get total number of active connections.

        Returns:
            Number of active connections
        """
        return len(self._clients)

    def get_document_count(self) -> int:
        """
        Get number of documents with active connections.

        Returns:
            Number of documents
        """
        return len(self._document_to_sockets)


# Global connection manager instance
connection_manager = ConnectionManager()
