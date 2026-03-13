"""
PPT Namespace

Handles PPT-specific Socket.IO events.

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


class PptNamespace(BaseNamespace):
    """
    PPT namespace (/ppt) for PowerPoint Add-In communication.

    Handles client → server events only:
    - ppt:event:slideChanged - Slide change notifications

    Server → client commands (ppt:get:*, ppt:insert:*, etc.) are sent via
    OfficeWorkspace.emit_to_document() using sio.call() for direct RPC.
    Results are returned via ack mechanism, no handlers needed here.
    """

    def __init__(self) -> None:
        super().__init__("/ppt")
        logger.info("PptNamespace initialized")

    # ========================================================================
    # Event Reporters (Client → Server events)
    # ========================================================================

    async def on_ppt_event_slideChanged(self, sid: str, data: Any) -> None:
        """
        Handle slide change event from PowerPoint Add-In.

        Event: ppt:event:slideChanged
        Direction: Client → Server (event report, fire-and-forget)

        Args:
            sid: Session ID
            data: {
                eventType: "slideChanged",
                clientId: str,
                documentUri: str,
                data: {
                    fromIndex: number,
                    toIndex: number
                },
                timestamp: number
            }
        """
        client_info = self.get_client_info(sid)
        if client_info:
            slide_data = data.get("data", {})
            from_index = slide_data.get("fromIndex", "?")
            to_index = slide_data.get("toIndex", "?")
            logger.info(f"PPT slide changed: {client_info.client_id}, from {from_index} to {to_index}")
