# filename: text_documents.py
# @Time    : 2024/4/29 14:36
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from enum import IntEnum
from typing import Annotated, Any, ClassVar

from annotated_types import Ge
from pydantic import BaseModel, ConfigDict, Field


class LSPPosition(BaseModel):
    """
    LSPPosition
    """

    # Line position in a document (zero-based).
    line: Annotated[int, Ge(0)]
    # Character offset on a line in a document (zero-based). The meaning of this offset is determined by the
    # negotiated `PositionEncodingKind`. If the character value is greater than the line length it defaults back
    # to the line length.
    character: Annotated[int, Ge(0)]


class LSPRange(BaseModel):
    """
    LSPRange
    """

    start: LSPPosition
    end: LSPPosition


class LSPTextDocumentItem(BaseModel):
    """
    LSPTextDocumentItem
    """

    # The text document's URI
    uri: str
    # The text document's language identifier.
    language_id: str = Field(validation_alias="languageId")
    # The version number of this document.
    version: int
    # The content of the opened text document.
    text: str


class LSPTextDocumentIdentifier(BaseModel):
    """
    TextDocumentIdentifier
    """

    # The text document's URI.
    uri: str


class LSPVersionedLSPTextDocumentIdentifier(LSPTextDocumentIdentifier):
    """
    LSPVersionedTextDocumentIdentifier
    """

    # The version number of this document.
    version: int


class LSPOptionalVersionedLSPTextDocumentIdentifier(LSPTextDocumentIdentifier):
    """
    LSPOptionalVersionedTextDocumentIdentifier
    """

    # The version number of this document.
    version: int | None = None


class LSPDocumentFilter(BaseModel):
    """
    LSPDocumentFilter
    """

    # A language id, like "python"
    language: str | None = None
    # A Uri [scheme](https://tools.ietf.org/html/rfc3986#section-3.1) like "file" or "untitled".
    f_schema: str | None = None  # schema是Pydantic预留字段
    # A glob pattern, like `*.{ts,js}`.
    pattern: str | None = None


class LSPTextEdit(BaseModel):
    """
    LSPTextEdit
    """

    range: LSPRange
    new_text: str = Field(validation_alias="newText")
    # 关键配置：允许同时使用字段名和别名进行赋值
    model_config: ClassVar[ConfigDict] = ConfigDict(populate_by_name=True)


class LSPChangeAnnotation(BaseModel):
    """
    LSPChangeAnnotation
    """

    label: str
    needs_confirmation: bool | None = Field(validation_alias="needsConfirmation")
    description: str | None = None


class LSPAnnotatedLSPTextEdit(LSPTextEdit):
    """
    LSPAnnotatedTextEdit
    """

    annotation_id: str = Field(validation_alias="annotationId")


class LSPTextDocumentEdit(BaseModel):
    """
    LSPTextDocumentEdit
    """

    text_document: LSPOptionalVersionedLSPTextDocumentIdentifier = Field(validation_alias="textDocument")


class LSPLocation(BaseModel):
    """
    LSPLocation
    """

    uri: str
    range: LSPRange


class DiagnosticSeverity(IntEnum):
    """
    Defines various levels of diagnostics.
    """

    Error = 1
    Warning = 2
    Information = 3
    Hint = 4


class DiagnosticTag(IntEnum):
    """
    Diagnostic tags are used to register quick fixes.  Each fix is registered on a code-action
    server and then later it can be converted into an LSP CodeAction.
    """

    Unnecessary = 1
    Deprecated = 2


class CodeDescription(BaseModel):
    """
    CodeDescription
    """

    href: str


class DiagnosticRelatedInformation(BaseModel):
    """
    DiagnosticRelatedInformation
    """

    location: LSPLocation
    message: str


class LSPDiagnostic(BaseModel):
    """
    LSPDiagnostic
    """

    # The range at which the message applies.
    range: LSPRange
    # The diagnostic's severity. Can be omitted. If omitted it is up to the client to interpret diagnostics as error,
    # warning, info or hint.
    severity: DiagnosticSeverity | None = None
    # The diagnostic's code. Can be omitted.
    code: int | str | None = None
    # An optional property to describe the error code.
    code_description: CodeDescription | None = Field(default=None, validation_alias="codeDescription")
    # A human-readable string describing the source of this diagnostic, e.g. 'typescript' or 'super lint'.
    source: str | None = None
    # The diagnostic's message.
    message: str
    # An array of related diagnostic information, e.g. when symbol-names within a scope collide all definitions can
    # have an diagnostic entries.
    related_information: list[DiagnosticRelatedInformation] | None = Field(validation_alias="relatedInformation")
    # Additional metadata about the diagnostic. 1 means Unused or unnecessary code. 2 means Deprecated or obsolete code.
    tags: list[DiagnosticTag] | None = None
    # A data entry field that is preserved between a `textDocument/publishDiagnostics` notification and
    # `textDocument/codeAction` request.
    data: Any = None
