"""
Word Namespace

Handles Word-specific Socket.IO events.

Architecture Note:
    Server → Client commands use OfficeWorkspace.emit_to_document() with sio.call(),
    which returns results directly via ack mechanism. No server-side handlers needed
    for command responses.

    This namespace only handles Client → Server event reports (fire-and-forget).
"""

import logging
from typing import Any

from .base import BaseNamespace

logger = logging.getLogger(__name__)


class WordNamespace(BaseNamespace):
    """
    Word namespace (/word) for Word Add-In communication.

    Handles client → server events only:
    - word:event:selectionChanged - Selection change notifications
    - word:event:documentModified - Document modification notifications

    Server → client commands (word:get:*, word:insert:*, etc.) are sent via
    OfficeWorkspace.emit_to_document() using sio.call() for direct RPC.
    Results are returned via ack mechanism, no handlers needed here.
    """

    def __init__(self) -> None:
        super().__init__("/word")
        logger.info("WordNamespace initialized")

    # ========================================================================
    # Event Reporters (Client → Server events)
    # ========================================================================

    async def on_word_event_selectionChanged(self, sid: str, data: Any) -> None:
        """
        Handle selection change event from Word Add-In.

        Event: word:event:selectionChanged
        Direction: Client → Server (event report, fire-and-forget)

        Args:
            sid: Session ID
            data: {
                eventType: "selectionChanged",
                clientId: str,
                documentUri: str,
                data: {
                    text: str,
                    length: number
                },
                timestamp: number
            }
        """
        client_info = self.get_client_info(sid)
        if client_info:
            logger.info(
                f"Word selection changed: {client_info.client_id}, text length: {data.get('data', {}).get('length', 0)}"
            )

    async def on_word_event_documentModified(self, sid: str, data: Any) -> None:
        """
        Handle document modification event from Word Add-In.

        Event: word:event:documentModified
        Direction: Client → Server (event report, fire-and-forget)

        Args:
            sid: Session ID
            data: {
                eventType: "documentModified",
                clientId: str,
                documentUri: str,
                data: {
                    modificationType: "insert" | "delete" | "update"
                },
                timestamp: number
            }
        """
        client_info = self.get_client_info(sid)
        if client_info:
            logger.info(
                f"Word document modified: {client_info.client_id}, type: {data.get('data', {}).get('modificationType')}"
            )
