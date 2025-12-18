# filename: file_resource.py
# @Time    : 2024/4/29 15:03
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from typing import Literal

from pydantic import BaseModel, Field


class LSPCreateFileOptions(BaseModel):
    """
    LSPCreateFileOptions
    """

    # Overwrite existing file.
    overwrite: bool | None = None
    # Ignore if file already exists.
    ignore_if_exists: bool | None = Field(None, validation_alias="ignoreIfExists")


class LSPCreateFile(BaseModel):
    """
    LSPCreateFile
    """

    kind: Literal["create"] = "create"
    # The resource to create.
    uri: str
    # Additional options
    options: LSPCreateFileOptions | None = None
    # An optional annotation identifier describing the operation.
    annotation_id: str | None = Field(None, validation_alias="annotationId")


class LSPRenameFileOptions(BaseModel):
    """
    LSPRenameFileOptions
    """

    # Overwrite target if existing. Overwrite wins over `ignoreIfExists`
    overwrite: bool | None = None
    # Ignores if target exists.
    ignore_if_exists: bool | None = Field(None, validation_alias="ignoreIfExists")


class LSPRenameFile(BaseModel):
    """
    LSPRenameFile
    """

    kind: Literal["rename"] = "rename"
    # The old (existing) location.
    old_uri: str = Field(..., validation_alias="oldUri")
    # The new location.
    new_uri: str = Field(..., validation_alias="newUri")
    # Additional options
    options: LSPRenameFileOptions | None = None
    # An optional annotation identifier describing the operation.
    annotation_id: str | None = Field(None, validation_alias="annotationId")


class LSPDeleteFileOptions(BaseModel):
    """
    LSPDeleteFileOptions
    """

    # Delete the content recursively if a folder is denoted.
    recursive: bool | None = None
    # Ignore the operation if the file doesn't exist.
    ignore_if_not_exists: bool | None = Field(None, validation_alias="ignoreIfNotExists")


class LSPDeleteFile(BaseModel):
    """
    LSPDeleteFile
    """

    kind: Literal["delete"] = "delete"
    # The file to delete.
    uri: str
    # Additional options
    options: LSPDeleteFileOptions | None = None
    # An optional annotation identifier describing the operation.
    annotation_id: str | None = Field(None, validation_alias="annotationId")
