# filename: commands.py
# @Time    : 2024/4/29 15:02
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from typing import Any

from pydantic import BaseModel


class LSPCommand(BaseModel):
    """
    Represents a reference to a command. Provides a title which will be used to represent a command in the UI. Commands
    are identified by a string identifier. The recommended way to handle commands is to implement their execution on
    the server side if the client and server provides the corresponding capabilities. Alternatively the tool extension
    code could handle the command. The protocol currently doesn’t specify a set of well-known commands.

    表示对命令的引用。提供一个标题，该标题将用于在 UI 中表示命令。命令由字符串标识符标识。处理命令的推荐方法是在服务器端实现其执行，或者如果
    客户端和服务器提供相应的功能，则在工具扩展代码中处理命令。LSP标准协议目前没有指定一组众所周知的命令。
    """

    # Title of the command, like "save".
    title: str
    # The identifier of the actual command handler.
    command: str
    # Arguments that the command handler should be invoked with.
    arguments: list[Any] | None = None
