"""
Connection Manager

Manages client connections and mappings between documentUri and socketId.
This is the core routing component for Workspace → Add-In communication.
"""

import logging
import os
import time
from collections.abc import Callable
from dataclasses import dataclass
from urllib.parse import unquote, urlparse

logger = logging.getLogger(__name__)


def normalize_document_uri(uri: str) -> str:
    """
    Normalize a document URI for consistent matching.

    Handles:
    1. Plain file paths (e.g., /Users/foo/doc.docx → file:///Users/foo/doc.docx)
    2. URL decoding (e.g., %2F → /)
    3. Symlink resolution (e.g., /var → /private/var on macOS)
    4. Path normalization

    Args:
        uri: Document URI (e.g., file:///path/to/doc.docx) or plain path (e.g., /path/to/doc.docx)

    Returns:
        Normalized file:// URI
    """
    if uri.startswith("file://"):
        try:
            # Parse the URI
            parsed = urlparse(uri)

            # URL-decode the path
            path = unquote(parsed.path)

            # On file:// URIs, the path might have an extra leading slash on some systems
            # e.g., file:///var/... or file:////var/... (Windows UNC paths)
            if path.startswith("//"):
                path = path[1:]
        except Exception as e:
            logger.warning(f"Failed to parse URI '{uri}': {e}")
            return uri
    elif uri.startswith("/"):
        # Plain absolute file path — normalize to file:// URI
        path = uri
    else:
        # Non-file URI (http://, etc.) — return as-is
        return uri

    # Resolve symlinks and normalize the path
    try:
        # Use os.path.realpath to resolve symlinks
        # This handles /var → /private/var on macOS
        real_path = os.path.realpath(path)
    except OSError:
        # If path doesn't exist, just normalize it
        real_path = os.path.normpath(path)

    # Reconstruct the URI
    return f"file://{real_path}"


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

        # Disconnect callbacks: called with (document_uri,) when a document has no more connections
        self._on_document_disconnect: list[Callable[[str], None]] = []

        # Connect callbacks: called with (document_uri, namespace) on first connection for a document
        self._on_document_connect: list[Callable[[str, str], None]] = []

        # Namespace-aware disconnect callbacks: called with (document_uri, namespace) on full disconnect
        self._on_document_disconnect_ns: list[Callable[[str, str], None]] = []

        logger.info("ConnectionManager initialized")

    def register_disconnect_callback(self, callback: Callable[[str], None]) -> None:
        """Register a callback invoked when a document loses all connections."""
        self._on_document_disconnect.append(callback)

    def register_connect_callback(self, callback: Callable[[str, str], None]) -> None:
        """Register a callback invoked with (document_uri, namespace) on first connection."""
        self._on_document_connect.append(callback)

    def register_disconnect_callback_ns(self, callback: Callable[[str, str], None]) -> None:
        """Register a callback invoked with (document_uri, namespace) when document fully disconnects."""
        self._on_document_disconnect_ns.append(callback)

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

        # Normalize document URI for consistent matching
        normalized_uri = normalize_document_uri(document_uri)
        if normalized_uri != document_uri:
            logger.debug(f"Normalized URI: {document_uri} → {normalized_uri}")

        client_info = ClientInfo(
            socket_id=socket_id,
            client_id=client_id,
            document_uri=normalized_uri,  # Store normalized URI
            namespace=namespace,
            connected_at=time.time(),
        )

        # Store client info
        self._clients[socket_id] = client_info
        self._client_id_to_socket[client_id] = socket_id

        # Map document_uri → socket_id (using normalized URI)
        is_first_connection = normalized_uri not in self._document_to_sockets
        if is_first_connection:
            self._document_to_sockets[normalized_uri] = set()
        self._document_to_sockets[normalized_uri].add(socket_id)

        logger.info(f"Client registered: {client_id} ({socket_id}) for {normalized_uri} on {namespace}")

        # Fire connect callbacks on first connection for this document
        if is_first_connection:
            for cb in self._on_document_connect:
                try:
                    cb(normalized_uri, namespace)
                except Exception:
                    logger.exception("Error in connect callback")

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
                # Notify listeners that this document is fully disconnected
                for cb in self._on_document_disconnect:
                    try:
                        cb(client_info.document_uri)
                    except Exception:
                        logger.exception("Error in disconnect callback")
                # Notify namespace-aware listeners
                for cb_ns in self._on_document_disconnect_ns:
                    try:
                        cb_ns(client_info.document_uri, client_info.namespace)
                    except Exception:
                        logger.exception("Error in disconnect_ns callback")

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
        # Normalize URI for lookup
        normalized_uri = normalize_document_uri(document_uri)
        sockets = self._document_to_sockets.get(normalized_uri)
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
        # Normalize URI for lookup
        normalized_uri = normalize_document_uri(document_uri)
        socket_ids = self._document_to_sockets.get(normalized_uri, set())
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
        # Normalize URI for lookup
        normalized_uri = normalize_document_uri(document_uri)
        return normalized_uri in self._document_to_sockets and bool(self._document_to_sockets[normalized_uri])

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
