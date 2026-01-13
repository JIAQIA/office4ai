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

        Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/29753356/word+insert+text

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri, text, location, format
                  format: {
                    bold?: boolean,
                    italic?: boolean,
                    fontSize?: number,
                    fontName?: string,
                    color?: string,
                    underline?: string,      // 🆕 Underline type
                    styleName?: string       // 🆕 Word style name
                  }

        Priority Rule (Important):
            - Direct formatting takes precedence over styleName
            - If any direct format fields (bold/italic/fontSize/etc) are provided,
              styleName is ignored
            - If only styleName is provided, Word style is applied
            - If neither is provided, default formatting is used

        Examples:
            # ❌ Not recommended: styleName will be ignored
            format = {"bold": True, "styleName": "Heading 1"}

            # ✅ Recommended: Use Word style only
            format = {"styleName": "Heading 1"}

            # ✅ Recommended: Use direct format only
            format = {"bold": True, "color": "#FF0000"}

        Error Codes:
            - 3001: DOCUMENT_NOT_FOUND - Document not found
            - 3004: INVALID_PARAM - Invalid parameter value
        """
        client_info = self.get_client_info(sid)
        if client_info:
            text = data.get("text", "")
            format_data = data.get("format", {})

            # Log format details including new fields
            log_parts = [
                f"text length: {len(text)}",
                f"location: {data.get('location', 'Cursor')}",
            ]

            if format_data:
                # Check if direct formatting is present
                has_direct_format = any(
                    format_data.get(key) is not None
                    for key in ["bold", "italic", "fontSize", "fontName", "color", "underline"]
                )
                has_style_name = format_data.get("styleName") is not None

                if has_direct_format:
                    direct_formats = []
                    if format_data.get("bold") is not None:
                        direct_formats.append(f"bold={format_data['bold']}")
                    if format_data.get("italic") is not None:
                        direct_formats.append(f"italic={format_data['italic']}")
                    if format_data.get("fontSize") is not None:
                        direct_formats.append(f"fontSize={format_data['fontSize']}")
                    if format_data.get("fontName") is not None:
                        direct_formats.append(f"fontName={format_data['fontName']}")
                    if format_data.get("color") is not None:
                        direct_formats.append(f"color={format_data['color']}")
                    if format_data.get("underline") is not None:
                        direct_formats.append(f"underline={format_data['underline']}")

                    log_parts.append(f"direct format: {', '.join(direct_formats)}")

                    # Log if styleName is present but will be ignored
                    if has_style_name:
                        log_parts.append(
                            f"⚠️  styleName '{format_data['styleName']}' will be ignored (direct format takes precedence)"
                        )
                elif has_style_name:
                    log_parts.append(f"styleName: {format_data['styleName']}")

            logger.info(
                f"Received word:insert:text from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}, "
                f"{', '.join(log_parts)}"
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
        """
        Handle word:get:documentStructure event from Add-In.

        Gets the document structure information, including paragraph count,
        table count, image count, and section count.

        Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri

        Response:
            {
                requestId: str,
                success: boolean,
                data: {
                    paragraphCount: number,
                    tableCount: number,
                    imageCount: number,
                    sectionCount: number
                },
                timestamp: number
            }

        Error Codes:
            - 3001: DOCUMENT_NOT_FOUND - Document not found

        Notes:
            - Returns document structure statistics, not actual content
            - All count values are integers representing total elements
            - Paragraph count includes empty paragraphs
            - Image count includes both inline and floating images
        """
        client_info = self.get_client_info(sid)
        if client_info:
            logger.info(
                f"Received word:get:documentStructure from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}"
            )

    async def on_word_get_documentStats(self, sid: str, data: Any) -> None:
        """
        Handle word:get:documentStats event from Add-In.

        Gets document statistics including word count, character count,
        and paragraph count.

        Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: Request data with requestId, documentUri

        Response:
            {
                requestId: str,
                success: boolean,
                data: {
                    wordCount: number,
                    characterCount: number,
                    paragraphCount: number
                },
                timestamp: number
            }

        Error Codes:
            - 3001: DOCUMENT_NOT_FOUND - Document not found

        Notes:
            - Word count uses Word's standard counting rules
            - Character count includes spaces and punctuation
            - Paragraph count includes empty paragraphs
            - Statistics cover entire document, not selection
        """
        client_info = self.get_client_info(sid)
        if client_info:
            logger.info(
                f"Received word:get:documentStats from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}"
            )

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
        """
        Handle word:replace:text event from Add-In.

        Find and replace text in Word document, with options for case sensitivity,
        whole word matching, and replacing all occurrences.

        Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30801921

        Note: This handler receives events from Add-In for logging/debugging.
        Server → Add-In commands should use OfficeWorkspace.emit_to_document().

        Args:
            sid: Session ID
            data: {
                requestId: str,
                documentUri: str,
                searchText: string,
                replaceText: string,
                options?: {
                    matchCase?: boolean,
                    matchWholeWord?: boolean,
                    replaceAll?: boolean
                },
                timestamp: number
            }

        Validation:
            - searchText must not be empty (MISSING_PARAM error 4001)
            - replaceText must not be empty (MISSING_PARAM error 4001)
            - matchCase: true = case sensitive, false = case insensitive
            - matchWholeWord: true = match whole words only
            - replaceAll: true = replace all, false = replace first only

        Response:
            {
                requestId: str,
                success: boolean,
                data?: {
                    replaceCount: number
                },
                error?: {
                    code: string,
                    message: string
                },
                timestamp: number
            }

        Error Codes:
            - 3001: DOCUMENT_NOT_FOUND - Document not found
            - 4001: MISSING_PARAM - searchText or replaceText is empty
        """
        client_info = self.get_client_info(sid)
        if client_info:
            search_text = data.get("searchText", "")
            replace_text = data.get("replaceText", "")
            options = data.get("options", {})

            # Validate required parameters
            if not search_text or not replace_text:
                logger.warning(
                    f"Invalid word:replace:text from {client_info.client_id}: "
                    f"searchText and replaceText required, requestId: {data.get('requestId', 'unknown')}"
                )
                return

            match_case = options.get("matchCase", False)
            match_whole_word = options.get("matchWholeWord", False)
            replace_all = options.get("replaceAll", False)

            logger.info(
                f"Received word:replace:text from {client_info.client_id}, "
                f"requestId: {data.get('requestId', 'unknown')}, "
                f"searchText: '{search_text[:50]}...', "
                f"replaceText: '{replace_text[:50]}...', "
                f"matchCase: {match_case}, "
                f"matchWholeWord: {match_whole_word}, "
                f"replaceAll: {replace_all}"
            )

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
