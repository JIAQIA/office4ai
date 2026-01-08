"""
Socket.IO Request Wrapper

Automatically wraps business parameters into BaseRequest-compatible DTOs
with auto-generated request IDs and timestamps.

This module now uses the auto-registration mechanism in BaseRequest:
- DTOs are registered automatically via __init_subclass__
- Global registry is managed by RequestRegistry singleton
- No manual registry maintenance required
"""

from typing import Any

from office4ai.environment.workspace.dtos.common import request_registry


class RequestWrapperError(Exception):
    """Raised when request wrapping fails"""

    pass


def wrap_request(
    event: str,
    business_params: dict[str, Any],
    document_uri: str | None = None,
) -> dict[str, Any]:
    """
    Wrap business parameters into a BaseRequest-compatible DTO.

    This function:
    1. Looks up the appropriate DTO class using the global request_registry
    2. Auto-generates a unique requestId (UUID4)
    3. Extracts document_uri from business_params if not provided
    4. Validates using Pydantic
    5. Returns camelCase JSON for Socket.IO transmission

    Note: This is now a thin wrapper around BaseRequest.from_event().to_payload().

    Args:
        event: Event name (e.g., "word:get:selectedContent")
        business_params: Business-specific parameters
        document_uri: Optional document URI (extracted from params if not provided)

    Returns:
        dict: JSON-ready dict with camelCase keys

    Raises:
        RequestWrapperError: If event not registered or validation fails

    Example:
        >>> params = {"options": {"includeText": True}}
        >>> wrapped = wrap_request("word:get:selectedContent", params, "file:///test.docx")
        >>> print(wrapped)
        {
            "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
            "documentUri": "file:///test.docx",
            "timestamp": 1234567890000,
            "options": {"includeText": True}
        }
    """
    from office4ai.environment.workspace.dtos.common import BaseRequest

    # Extract document_uri
    if not document_uri:
        document_uri = business_params.get("document_uri")
        if not document_uri:
            raise RequestWrapperError("document_uri must be provided either as argument or in business_params")

    # Remove document_uri from business_params to avoid duplicate argument
    business_params_copy = {**business_params}
    business_params_copy.pop("document_uri", None)

    # Use BaseRequest.from_event() to create the request
    try:
        request = BaseRequest.from_event(event, document_uri, **business_params_copy)
        return request.to_payload()
    except RequestWrapperError:
        # Re-raise RequestWrapperError as-is
        raise
    except Exception as e:
        # Wrap other exceptions
        raise RequestWrapperError(f"Failed to wrap request for event '{event}': {e}") from e


def is_wrappable_event(event: str) -> bool:
    """
    Check if an event has a registered DTO wrapper.

    Args:
        event: Event name

    Returns:
        True if event is registered, False otherwise
    """
    return request_registry.contains(event)


def get_registered_events() -> list[str]:
    """
    Get list of all registered event names.

    Returns:
        Sorted list of event names
    """
    return request_registry.all_events()
