"""Shared fixtures for resource unit tests."""

from __future__ import annotations

import time

import pytest

from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import ClientInfo


@pytest.fixture
def workspace() -> OfficeWorkspace:
    """Create a lightweight OfficeWorkspace without starting Socket.IO."""
    ws = OfficeWorkspace.__new__(OfficeWorkspace)
    ws._last_activity = None
    ws._content_cache = {}
    ws._structure_cache = {}
    return ws


def make_client(
    doc_uri: str,
    namespace: str = "/word",
    socket_id: str = "s1",
    client_id: str = "c1",
) -> ClientInfo:
    """Create a ClientInfo for testing."""
    return ClientInfo(
        socket_id=socket_id,
        client_id=client_id,
        document_uri=doc_uri,
        namespace=namespace,
        connected_at=time.time(),
    )
