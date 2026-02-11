# filename: tools.py
# @Time    : 2025/10/29 12:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
MCP 工具的 Schema 定义 | MCP Tools Schema Definitions

定义所有工具的输入输出 Schema
Defines input/output schemas for all tools
"""

from typing import Any

from pydantic import BaseModel, Field


class BashInput(BaseModel):
    """Bash 工具输入 Schema | Bash Tool Input Schema"""

    command: str = Field(..., description="要执行的命令 | The command to execute")
    args: str | list[str] | None = Field(
        default=None,
        description="命令参数，可以是字符串或字符串列表 | Command arguments, can be a string or list of strings",
    )
    timeout: int | None = Field(
        default=None,
        description="超时时间(毫秒)，最大 600000 | Optional timeout in milliseconds (max 600000)",
        le=600000,
    )
    description: str | None = Field(
        default=None,
        description="命令的简洁描述(5-10个词) | Clear, concise description of what this command does in 5-10 words",
    )
    run_in_background: bool = Field(
        default=False,
        description="是否在后台运行 | Set to true to run this command in the background",
    )
    dangerously_disable_sandbox: bool = Field(
        default=False,
        description="危险：禁用沙箱模式 | Set this to true to dangerously override sandbox mode",
    )


class BashOutput(BaseModel):
    """Bash 工具输出 Schema | Bash Tool Output Schema"""

    success: bool = Field(..., description="命令是否成功执行 | Whether the command executed successfully")
    output: str = Field(..., description="命令输出 | Command output")
    error: str | None = Field(default=None, description="错误信息(如果有) | Error message (if any)")
    exit_code: int | None = Field(default=None, description="退出码 | Exit code")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据 | Additional metadata",
    )


class GlobInput(BaseModel):
    """Glob 工具输入 Schema | Glob Tool Input Schema"""

    pattern: str = Field(..., description="Glob 模式 | The glob pattern to match files against")
    path: str | None = Field(
        default=None,
        description="搜索目录，默认为当前工作目录 | The directory to search in. Defaults to current working directory",
    )


class GlobOutput(BaseModel):
    """Glob 工具输出 Schema | Glob Tool Output Schema"""

    success: bool = Field(..., description="是否成功执行 | Whether the operation was successful")
    files: list[dict[str, Any]] = Field(
        default_factory=list,
        description="匹配的文件列表 | List of matched files with path and metadata",
    )
    error: str | None = Field(default=None, description="错误信息(如果有) | Error message (if any)")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据 | Additional metadata",
    )


class GrepInput(BaseModel):
    """Grep 工具输入 Schema | Grep Tool Input Schema"""

    pattern: str = Field(..., description="正则表达式模式 | The regular expression pattern to search for")
    path: str | None = Field(default=None, description="搜索路径，默认为当前工作目录 | File or directory to search in")
    glob: str | None = Field(default=None, description="文件过滤 Glob 模式 | Glob pattern to filter files")
    output_mode: str = Field(
        default="files_with_matches",
        description="输出模式 | Output mode: content, files_with_matches, or count",
    )
    context_before: int | None = Field(default=None, alias="-B", description="显示匹配前的行数 | Lines before match")
    context_after: int | None = Field(default=None, alias="-A", description="显示匹配后的行数 | Lines after match")
    context: int | None = Field(
        default=None, alias="-C", description="显示匹配前后的行数 | Lines before and after match"
    )
    line_number: bool | None = Field(default=None, alias="-n", description="显示行号 | Show line numbers")
    case_insensitive: bool | None = Field(default=None, alias="-i", description="忽略大小写 | Case insensitive search")
    file_type: str | None = Field(default=None, alias="type", description="文件类型 | File type to search")
    head_limit: int | None = Field(default=None, description="限制输出行数 | Limit output to first N lines")
    multiline: bool = Field(default=False, description="启用多行模式 | Enable multiline mode")


class GrepOutput(BaseModel):
    """Grep 工具输出 Schema | Grep Tool Output Schema"""

    success: bool = Field(..., description="是否成功执行 | Whether the operation was successful")
    output: str = Field(default="", description="搜索输出结果 | Search output result")
    matched: bool = Field(default=False, description="是否找到匹配 | Whether matches were found")
    error: str | None = Field(default=None, description="错误信息(如果有) | Error message (if any)")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据 | Additional metadata",
    )


class ReadInput(BaseModel):
    """Read 工具输入 Schema | Read Tool Input Schema"""

    file_path: str = Field(..., description="要读取的文件的绝对路径 | The absolute path to the file to read")
    offset: int | None = Field(default=None, description="起始行号 | The line number to start reading from", ge=1)
    limit: int | None = Field(default=None, description="读取的行数 | The number of lines to read", ge=1)


class ReadOutput(BaseModel):
    """Read 工具输出 Schema | Read Tool Output Schema"""

    success: bool = Field(..., description="是否成功执行 | Whether the operation was successful")
    content: str = Field(default="", description="文件内容 | File content")
    error: str | None = Field(default=None, description="错误信息(如果有) | Error message (if any)")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据 | Additional metadata",
    )


class EditInput(BaseModel):
    """Edit 工具输入 Schema | Edit Tool Input Schema"""

    file_path: str = Field(..., description="要修改的文件的绝对路径 | The absolute path to the file to modify")
    old_string: str = Field(..., description="要替换的文本 | The text to replace")
    new_string: str = Field(
        ...,
        description="替换后的文本（必须与 old_string 不同）| The text to replace it with (must be different from old_string)",
    )
    replace_all: bool = Field(
        default=False, description="是否替换所有匹配项（默认 false）| Replace all occurrences (default false)"
    )


class EditOutput(BaseModel):
    """Edit 工具输出 Schema | Edit Tool Output Schema"""

    success: bool = Field(..., description="是否成功执行 | Whether the operation was successful")
    message: str = Field(default="", description="操作结果消息 | Operation result message")
    replacements_made: int = Field(default=0, description="执行的替换次数 | Number of replacements made")
    error: str | None = Field(default=None, description="错误信息(如果有) | Error message (if any)")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据 | Additional metadata",
    )


class WriteInput(BaseModel):
    """Write 工具输入 Schema | Write Tool Input Schema"""

    file_path: str = Field(..., description="要写入的文件的绝对路径 | The absolute path to the file to write")
    content: str = Field(..., description="要写入的内容 | The content to write to the file")


class WriteOutput(BaseModel):
    """Write 工具输出 Schema | Write Tool Output Schema"""

    success: bool = Field(..., description="是否成功执行 | Whether the operation was successful")
    message: str = Field(default="", description="操作结果消息 | Operation result message")
    error: str | None = Field(default=None, description="错误信息(如果有) | Error message (if any)")
    metadata: dict[str, Any] = Field(
        default_factory=dict,
        description="额外的元数据 | Additional metadata",
    )
