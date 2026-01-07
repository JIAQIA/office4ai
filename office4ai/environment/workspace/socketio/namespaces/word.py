"""
Word Namespace

Handles Word-specific Socket.IO events.
"""

import logging
from typing import Any

from ..request_handler import emit_with_response, handle_response
from .base import BaseNamespace

logger = logging.getLogger(__name__)


class WordNamespace(BaseNamespace):
    """
    Word namespace (/word) for Word Add-In communication.

    Handles events:
    - word:get:selectedContent
    - word:insert:text
    - word:replace:selection

    MVP: Implements 3 core events
    Future: Will implement all 13 Word events
    """

    def __init__(self) -> None:
        super().__init__("/word")
        logger.info("WordNamespace initialized")

    # ========================================================================
    # Event Handlers (Server → Client commands)
    # ========================================================================

    async def on_word_get_selectedContent(self, sid: str, data: Any) -> None:
        """
        Handle request to get selected content from Word.

        Event: word:get:selectedContent
        Direction: Server → Client (command)
        Response: word:get:selectedContent:response

        Args:
            sid: Session ID
            data: {
                requestId: str,
                documentUri: str,
                options?: GetContentOptions
            }
        """
        logger.debug(f"word:get:selectedContent from {sid}: {data}")

        # Forward request to Add-In using request-response mechanism
        # The Add-In will execute the Word API call and send back response
        try:
            response = await emit_with_response(sid=sid, event="word:get:selectedContent", data=data, timeout=10.0)
            logger.info(f"Got selected content response: {response}")
        except Exception as e:
            logger.error(f"Error getting selected content: {e}")

    async def on_word_insert_text(self, sid: str, data: Any) -> None:
        """
        Handle request to insert text into Word.

        Event: word:insert:text
        Direction: Server → Client (command)
        Response: word:insert:text:response

        Args:
            sid: Session ID
            data: {
                requestId: str,
                documentUri: str,
                text: str,
                location?: "Cursor" | "Start" | "End",
                format?: TextFormat
            }
        """
        logger.debug(f"word:insert:text from {sid}: {data}")

        # Forward request to Add-In using request-response mechanism
        try:
            response = await emit_with_response(sid=sid, event="word:insert:text", data=data, timeout=10.0)
            logger.info(f"Insert text response: {response}")
        except Exception as e:
            logger.error(f"Error inserting text: {e}")

    async def on_word_replace_selection(self, sid: str, data: Any) -> None:
        """
        Handle request to replace selected content in Word.

        Event: word:replace:selection
        Direction: Server → Client (command)
        Response: word:replace:selection:response

        Args:
            sid: Session ID
            data: {
                requestId: str,
                documentUri: str,
                content: {
                    text?: str,
                    images?: ImageData[],
                    format?: TextFormat
                }
            }
        """
        logger.debug(f"word:replace:selection from {sid}: {data}")

        # Forward request to Add-In using request-response mechanism
        try:
            response = await emit_with_response(sid=sid, event="word:replace:selection", data=data, timeout=10.0)
            logger.info(f"Replace selection response: {response}")
        except Exception as e:
            logger.error(f"Error replacing selection: {e}")

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
    # Response Handlers (Client → Server responses)
    # ========================================================================

    async def on_word_get_selectedContent_response(self, sid: str, data: Any) -> None:
        """
        Handle response from Add-In for get selected content request.

        Event: word:get:selectedContent:response
        Direction: Client → Server (response)

        Args:
            sid: Session ID
            data: {
                requestId: str,
                success: boolean,
                content?: {
                    text: str,
                    html?: string,
                    images?: ImageData[]
                },
                error?: string
            }
        """
        request_id = data.get("requestId")
        if request_id:
            logger.debug(f"Received response for requestId={request_id}")
            handle_response(request_id, data)

    async def on_word_insert_text_response(self, sid: str, data: Any) -> None:
        """
        Handle response from Add-In for insert text request.

        Event: word:insert:text:response
        Direction: Client → Server (response)

        Args:
            sid: Session ID
            data: {
                requestId: str,
                success: boolean,
                error?: string
            }
        """
        request_id = data.get("requestId")
        if request_id:
            logger.debug(f"Received response for requestId={request_id}")
            handle_response(request_id, data)

    async def on_word_replace_selection_response(self, sid: str, data: Any) -> None:
        """
        Handle response from Add-In for replace selection request.

        Event: word:replace:selection:response
        Direction: Client → Server (response)

        Args:
            sid: Session ID
            data: {
                requestId: str,
                success: boolean,
                error?: string
            }
        """
        request_id = data.get("requestId")
        if request_id:
            logger.debug(f"Received response for requestId={request_id}")
            handle_response(request_id, data)

    # ========================================================================
    # Future Events (to be implemented)
    # ========================================================================

    # Content retrieval
    async def on_word_get_visibleContent(self, sid: str, data: Any) -> None:
        """TODO: Get visible content"""
        logger.warning("word:get:visibleContent not yet implemented")

    async def on_word_get_documentStructure(self, sid: str, data: Any) -> None:
        """TODO: Get document structure"""
        logger.warning("word:get:documentStructure not yet implemented")

    async def on_word_get_documentStats(self, sid: str, data: Any) -> None:
        """TODO: Get document statistics"""
        logger.warning("word:get:documentStats not yet implemented")

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
