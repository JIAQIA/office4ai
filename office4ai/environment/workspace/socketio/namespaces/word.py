"""
Word Namespace

Handles Word-specific Socket.IO events.
"""

import logging
from typing import Any

from .base import BaseNamespace

logger = logging.getLogger(__name__)


class WordNamespace(BaseNamespace):
    """
    Word namespace (/word) for Word Add-In communication.

    Handles client → server events:
    - word:event:selectionChanged - Selection change notifications
    - word:event:documentModified - Document modification notifications

    Server → client commands should be sent via OfficeWorkspace.emit_to_document()
    which uses sio_server.call() for direct RPC with automatic acknowledgement.
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

    # ========================================================================
    # Request Handlers (Log Only - commands sent via OfficeWorkspace.emit_to_document())
    # ========================================================================

    async def on_word_get_selectedContent(self, sid: str, data: Any) -> None:
        """
        Handle word:get:selectedContent event from Add-In.

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri, options
        """
        client_info = self.get_client_info(sid)
        if client_info:
            logger.info(
                f"Received word:get:selectedContent from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}"
            )

    async def on_word_insert_text(self, sid: str, data: Any) -> None:
        """
        Handle word:insert:text event from Add-In.

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri, text, location
        """
        client_info = self.get_client_info(sid)
        if client_info:
            logger.info(
                f"Received word:insert:text from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}, "
                f"text length: {len(data.get('text', ''))}"
            )

    async def on_word_replace_selection(self, sid: str, data: Any) -> None:
        """
        Handle word:replace:selection event from Add-In.

        Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30605313

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: {
                requestId: str,
                documentUri: str,
                content: {
                    text?: string,
                    images?: ImageData[],
                    format?: TextFormat
                },
                timestamp: number
            }

        Validation:
            - content.text or content.images must be provided
            - format only applies to text content
            - Replaces entire selection, original formatting not preserved

        Error Codes:
            - 3001: DOCUMENT_NOT_FOUND - Document not found
            - 3002: SELECTION_EMPTY - Current selection is empty
            - 3003: DOCUMENT_READ_ONLY - Document is read-only
        """
        client_info = self.get_client_info(sid)
        if client_info:
            content = data.get("content", {})
            text = content.get("text")
            images = content.get("images")

            # Validate that at least text or images is provided
            if not text and not images:
                logger.warning(
                    f"Invalid word:replace:selection from {client_info.client_id}: "
                    f"content.text or content.images required, requestId: {data.get('requestId', 'unknown')}"
                )
                return

            character_count = len(text) if text else 0

            logger.info(
                f"Received word:replace:selection from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}, "
                f"text length: {character_count}, "
                f"images: {len(images) if images else 0}"
            )

    # ========================================================================
    # Future Events (to be implemented)
    # ========================================================================

    # Content retrieval
    async def on_word_get_visibleContent(self, sid: str, data: Any) -> None:
        """
        Handle word:get:visibleContent event from Add-In.

        Gets the visible content from the current view, including text,
        images, tables, and other elements.

        Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30736386

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri, options
                  options: {
                    includeText?: boolean,
                    includeImages?: boolean,
                    includeTables?: boolean,
                    maxTextLength?: number
                  }

        Error Codes:
            - 3001: DOCUMENT_NOT_FOUND - Document not found
        """
        client_info = self.get_client_info(sid)
        if client_info:
            options = data.get("options", {})
            logger.info(
                f"Received word:get:visibleContent from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}, "
                f"includeText: {options.get('includeText', True)}, "
                f"includeImages: {options.get('includeImages', True)}, "
                f"includeTables: {options.get('includeTables', True)}"
            )

    async def on_word_get_documentStructure(self, sid: str, data: Any) -> None:
        """TODO: Get document structure"""
        logger.warning("word:get:documentStructure not yet implemented")

    async def on_word_get_documentStats(self, sid: str, data: Any) -> None:
        """TODO: Get document statistics"""
        logger.warning("word:get:documentStats not yet implemented")

    async def on_word_get_styles(self, sid: str, data: Any) -> None:
        """
        Handle word:get:styles event from Add-In.

        Gets all available styles from the Word document, including built-in
        and custom styles, with filtering options.

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri, options
                  options: {
                    includeBuiltIn: boolean,
                    includeCustom: boolean,
                    includeUnused: boolean,
                    detailedInfo: boolean
                  }
        """
        client_info = self.get_client_info(sid)
        if client_info:
            options = data.get("options", {})
            logger.info(
                f"Received word:get:styles from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}, "
                f"includeBuiltIn: {options.get('includeBuiltIn', True)}, "
                f"includeCustom: {options.get('includeCustom', True)}"
            )

    # Text operations
    async def on_word_replace_text(self, sid: str, data: Any) -> None:
        """TODO: Find and replace text"""
        logger.warning("word:replace:text not yet implemented")

    async def on_word_append_text(self, sid: str, data: Any) -> None:
        """TODO: Append text"""
        logger.warning("word:append:text not yet implemented")

    # Multimedia operations
    async def on_word_insert_image(self, sid: str, data: Any) -> None:
        """TODO: Insert image"""
        logger.warning("word:insert:image not yet implemented")

    async def on_word_insert_table(self, sid: str, data: Any) -> None:
        """TODO: Insert table"""
        logger.warning("word:insert:table not yet implemented")

    async def on_word_insert_equation(self, sid: str, data: Any) -> None:
        """TODO: Insert equation"""
        logger.warning("word:insert:equation not yet implemented")

    # Advanced features
    async def on_word_insert_toc(self, sid: str, data: Any) -> None:
        """TODO: Insert table of contents"""
        logger.warning("word:insert:toc not yet implemented")

    async def on_word_export_content(self, sid: str, data: Any) -> None:
        """TODO: Export content"""
        logger.warning("word:export:content not yet implemented")
