"""AI IDE Exceptions Module."""


class IDEExecutionError(Exception):
    """
    Exception raised for errors during execution of IDE tasks.

    Attributes:
        message: Error message
        detail_for_llm: Detailed error message for LLM
        call_id: Call identifier
        tool_name: Name of the tool that raised the error
        success: Whether the operation was successful
        done: Whether the operation is done
    """

    def __init__(
        self,
        message: str,
        detail_for_llm: str | None = None,
        call_id: str | int | None = None,
        tool_name: str | None = None,
        success: bool = False,
        done: bool = False,
    ) -> None:
        self.message = message
        self.detail_for_llm = detail_for_llm
        self.call_id = call_id
        self.tool_name = tool_name
        self.success = success
        self.done = done
        super().__init__(message)


class IDEProtocolError(Exception):
    """
    Exception raised for errors in the IDE protocol.

    Attributes:
        message: Error message
        protocol: Protocol type that caused the error
    """

    def __init__(self, message: str, protocol: str = "unknown") -> None:
        self.message = message
        self.protocol = protocol
        super().__init__(message)
