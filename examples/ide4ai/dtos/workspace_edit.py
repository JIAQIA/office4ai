# filename: workspace_edit.py
# @Time    : 2024/4/29 15:23
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from pydantic import BaseModel, Field

from ide4ai.dtos.file_resource import (
    LSPCreateFile,
    LSPDeleteFile,
    LSPRenameFile,
)
from ide4ai.dtos.text_documents import (
    LSPChangeAnnotation,
    LSPTextDocumentEdit,
    LSPTextEdit,
)


class LSPWorkspaceEdit(BaseModel):
    """
    A workspace edit represents changes to many resources managed in the workspace. The edit should either provide
    changes or documentChanges. If the client can handle versioned document edits and if documentChanges are present,
    the latter are preferred over changes.
    """

    changes: dict[str, list[LSPTextEdit]] | None = None

    documentChanges: (
        list[LSPTextDocumentEdit] | list[LSPTextDocumentEdit | LSPCreateFile | LSPRenameFile | LSPDeleteFile] | None
    ) = None

    change_annotations: dict[str, LSPChangeAnnotation] | None = Field(
        default=None,
        validation_alias="changeAnnotations",
    )
