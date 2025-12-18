# filename: utils.py
# @Time    : 2024/4/22 12:23
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

import os
import re
import unicodedata
from pathlib import Path
from re import Pattern

from ide4ai.environment.workspace.schema import EndOfLineSequence


def detect_newline_type(uri: Path) -> EndOfLineSequence:
    """
    Detects the newline character in a file, accounting for cases where the first newline
    may occur after an initial large block of data.

    If the file is empty or one-line, the default newline character is returned based on the operating system.

    Args:
        uri (Path): Path to the file.

    Returns:
        EndOfLineSequence: A string enum the newline type: '\\n', '\\r\\n'.
    """
    try:
        with open(uri, "rb") as file:
            while True:
                content = file.read(1024)
                if not content:  # 文件结束
                    break
                if b"\r\n" in content:
                    return EndOfLineSequence.CRLF
                elif b"\n" in content:
                    return EndOfLineSequence.LF
    except OSError as e:
        raise ValueError(f"Error reading file {uri}: {e}") from e

    # Determine the default newline based on the operating system
    if os.name == "nt":  # Windows
        return EndOfLineSequence.CRLF
    elif os.name == "posix":
        if os.uname().sysname == "Darwin":  # macOS
            return EndOfLineSequence.LF
        else:  # Linux and other POSIX systems
            return EndOfLineSequence.LF
    else:
        raise ValueError("Unsupported operating system")


LARGE_FILE_SIZE_THRESHOLD = 100 * 1024  # 100KB
LARGE_FILE_LINE_COUNT_THRESHOLD = 3_000  # 3000 lines
LARGE_FILE_HEAP_OPERATION_THRESHOLD = 256 * 1024  # 256K characters, usually ~> 512KB memory usage
UTF_8_BOM = "\ufeff"


def read_file_with_bom_handling(path: str) -> tuple[str, list[str]]:
    # 检查文件大小
    file_size = os.path.getsize(path)
    if file_size > LARGE_FILE_SIZE_THRESHOLD:
        raise ValueError("File size exceeds the maximum limit of 100KB.")

    try:
        with open(path, encoding="utf-8", newline="") as f:
            content: list[str] = f.read().splitlines()
            if not content:
                content = [""]  # pragma: no cover

        # 检查行数
        if len(content) > LARGE_FILE_LINE_COUNT_THRESHOLD:
            raise ValueError("File has more lines than the maximum allowed 300 lines.")

        # 检查字符总数，以评估内存使用
        total_characters = sum(len(line) for line in content)
        if total_characters > LARGE_FILE_HEAP_OPERATION_THRESHOLD:
            raise ValueError("File content exceeds the memory usage threshold of 256K characters.")

        if len(content) > 0 and content[0].startswith(UTF_8_BOM):
            bom = UTF_8_BOM
            content[0] = content[0].lstrip(UTF_8_BOM)
        else:
            bom = ""

        return bom, content
    except OSError as e:
        raise ValueError(f"Error reading file {path}: {e}") from e  # pragma: no cover
    except UnicodeDecodeError as e:
        raise ValueError("Only UTF-8 encoded files are supported now.") from e


def is_high_surrogate(char: str) -> bool:
    """判断给定字符是否是高代理（high surrogate）。

    Args:
        char (str): 输入的字符。

    Returns:
        bool: 如果是高代理字符返回 True，否则返回 False。
    """
    # 获取字符的 Unicode 码点
    char_code = ord(char)
    # 检查码点是否在高代理范围内
    return 0xD800 <= char_code <= 0xDBFF


def next_render_tab_stop(visible_column: int, tab_size: int) -> int:
    """
    Args:
        visible_column (int): The current visible column position.
        tab_size (int): The size of tabs in terms of number of spaces.

    Returns:
        int: The next column position that should be considered as a tab stop.
    """
    return visible_column + tab_size - visible_column % tab_size


def next_indent_tab_stop(visible_column: int, indent_size: int) -> int:
    """
    Args:
        visible_column (int): The current visible column.
        indent_size (int): The size of each indent.

    Returns:
        int: The next tab stop column after the current visible column.
    """
    return visible_column + indent_size - visible_column % indent_size


def prev_render_tab_stop(column: int, tab_size: int) -> int:
    """
    Args:
        column (int): The current column position.
        tab_size (int): The size of each tab stop.

    Returns:
        int: The column position of the previous tab stop.

    """
    return max(0, column - 1 - (column - 1) % tab_size)


def prev_indent_tab_stop(column: int, indent_size: int) -> int:
    """
    Args:
        column: An integer representing the current column position.
        indent_size: An integer representing the size of each indentation.

    Returns:
        An integer representing the column position of the previous indentation tab stop.

    """
    return max(0, column - 1 - (column - 1) % indent_size)


def is_full_width_character(code_point: int) -> bool:
    """
    Args:
        code_point: An integer value representing the Unicode code point of a character.

    Returns:
        A boolean value indicating whether the character at the given code point is a full-width character or not.
        Returns True if the character is full-width, and False otherwise.
    """
    return unicodedata.east_asian_width(chr(code_point)) in "FW"


def is_emoji_imprecise(code_point: int) -> bool:
    """
    Args:
        code_point (int): The code point of the character to check.

    Returns:
        bool: True if the character is considered an imprecise emoji, False otherwise.
    """
    try:
        import emoji

        character = chr(code_point)
        return emoji.is_emoji(character)
    except ImportError:  # pragma: no cover
        # 回退到使用unicodedata库的逻辑
        try:  # pragma: no cover
            return "EMOJI" in unicodedata.name(chr(code_point), "")  # pragma: no cover
        except ValueError:  # pragma: no cover
            return False  # 无效的Unicode码点  # pragma: no cover


def _next_visible_column(code_point: int, visible_column: int, tab_size: int) -> int:
    """
    Args:
        code_point: An integer representing the code point of the character.
        visible_column: An integer representing the current visible column.
        tab_size: An integer representing the size of a tab in number of spaces.

    Returns:
        An integer representing the next visible column after processing the given code point.

    """
    if code_point == 9:  # CharCode.Tab
        return next_render_tab_stop(visible_column, tab_size)
    if is_full_width_character(code_point) or is_emoji_imprecise(code_point):
        return visible_column + 2
    return visible_column + 1


def visible_width_from_column(line_content: str, column: int, tab_size: int) -> int:
    """
    注意：这是计算 visible_width 用于渲染与展示用，而不是计算光标位置使用，在全宽字符下（比如中文） column偏移是1，但visible_width是2

    Args:
        line_content: A string representing the content of the line.
        column: An integer representing the column number.
        tab_size: An integer representing the number of spaces to consider for a tab.

    Returns:
        An integer representing the visible column number.

    """
    text_len = min(column - 1, len(line_content))
    text = line_content[:text_len]
    result = 0
    offset = 0
    while offset < len(text):
        code_point = ord(text[offset])
        result = _next_visible_column(code_point, result, tab_size)
        offset += 1
    return result


def column_from_visible_width(line_content: str, visible_column: int, tab_size: int) -> int:
    """
    Args:
        line_content: A string representing the content of a line.
        visible_column: An integer representing the target visible column number.
        tab_size: An integer representing the number of spaces occupied by a tab.

    Returns:
        An integer representing the column number corresponding to the given visible column.

    """
    if visible_column <= 0:
        return 1
    before_visible_column = 0
    before_column = 1
    offset = 0
    while offset < len(line_content):
        code_point = ord(line_content[offset])
        after_visible_column = _next_visible_column(code_point, before_visible_column, tab_size)
        after_column = offset + 1 + 1
        if after_visible_column >= visible_column:
            before_delta = visible_column - before_visible_column
            after_delta = after_visible_column - visible_column
            if after_delta < before_delta:
                return after_column
            else:
                return before_column
        before_visible_column = after_visible_column
        before_column = after_column
        offset += 1
    return len(line_content) + 1


def _normalize_indentation_from_whitespace(s: str, indent_size: int, insert_spaces: bool) -> str:
    """
    Args:
        s: The input string.
        indent_size: The size of an indent, either in spaces or tabs.
        insert_spaces: A boolean flag indicating whether to insert spaces (True) or tabs (False) for indentation.

    Returns:
        A new string with normalized indentation based on the provided parameters.

    """
    spaces_cnt = 0
    for char in s:
        if char == "\t":
            spaces_cnt = next_indent_tab_stop(spaces_cnt, indent_size)
        else:
            spaces_cnt += 1
    result = ""
    if not insert_spaces:
        tabs_cnt = spaces_cnt // indent_size
        spaces_cnt %= indent_size
        result += "\t" * tabs_cnt
    result += " " * spaces_cnt
    return result


def normalize_indentation(s: str, indent_size: int, insert_spaces: bool) -> str:
    """
    主要解决空格与制表符混排的问题，如果insert_spaces设置为True，则最终会将制表符替换为空格。否则会以制表符点位，剩余的位置使用空格补齐

    Args:
        s: A string representing the input text that needs to be normalized.
        indent_size: An integer representing the desired size of each indentation level.
        insert_spaces: A boolean indicating whether to insert spaces or tabs for indentation.

    Returns:
        A string with normalized indentation. Each indentation level in the input text is adjusted to have the specified
         size and uses the specified type of indentation (spaces or tabs).

    Examples:
        >>> normalize_indentation("    hello", 4, True)
        >>> # returns "    hello"  # Already normalized with spaces
        >>> normalize_indentation("\thello", 4, True)
        >>> # returns "    hello"  # Convert tabs to spaces
        >>> normalize_indentation("    hello", 4, False)
        >>> # returns "\thello"  # Convert spaces to tabs
        >>> normalize_indentation("\t\thello", 2, True)
        >>> # returns "    hello"  # Convert larger tab stops to spaces
        >>> normalize_indentation("  \t  hello", 4, False)
        >>> # returns "\t  hello"  # Mixed spaces and tabs to tabs
        >>> normalize_indentation("hello", 4, True)
        >>> # returns "hello"  # No leading whitespace
        >>> normalize_indentation("\t\t  hello", 4, False)
        >>> # returns "\t\t  hello"  # Mixed, with tabs staying and spaces converting
    """
    first_non_whitespace_column = next((i for i, char in enumerate(s) if not char.isspace()), len(s))
    return (
        _normalize_indentation_from_whitespace(s[:first_non_whitespace_column], indent_size, insert_spaces)
        + s[first_non_whitespace_column:]
    )


def first_non_whitespace_index(s: str) -> int:
    """
    Returns the first index of the string that is not whitespace.
    If the string is empty or contains only whitespaces, returns -1.
    """
    for i, char in enumerate(s):
        if not char.isspace():
            return i
    return -1


def last_non_whitespace_index(s: str, start_index: int | None = None) -> int:
    """
    Returns the last index of the string that is not whitespace.
    If the string is empty or contains only whitespaces, or the start index is out of range, returns -1.
    """
    if start_index is None:
        start_index = len(s) - 1
    start_index = min(start_index, len(s) - 1)
    for i in range(start_index, -1, -1):
        if not s[i].isspace():
            return i
    return -1


def make_contains_rtl() -> Pattern:
    # Generated using https://github.com/alexdima/unicode-utils/blob/main/rtl-test.js
    return re.compile(
        r"(?:[\u05BE\u05C0\u05C3\u05C6\u05D0-\u05F4\u0608\u060B\u060D\u061B-\u064A\u066D-\u066F\u0671-\u06D5\u06E5"  # noqa
        r"\u06E6\u06EE\u06EF\u06FA-\u0710\u0712-\u072F\u074D-\u07A5\u07B1-\u07EA\u07F4\u07F5\u07FA\u07FE-\u0815\u081A"
        r"\u0824\u0828\u0830-\u0858\u085E-\u088E\u08A0-\u08C9\u200F\uFB1D\uFB1F-\uFB28\uFB2A-\uFD3D\uFD50-\uFDC7"
        r"\uFDF0-\uFDFC\uFE70-\uFEFC]"
        r"|\uD802[\uDC00-\uDD1B\uDD20-\uDE00\uDE10-\uDE35\uDE40-\uDEE4\uDEEB-\uDF35\uDF40-\uDFFF]"
        r"|\uD803[\uDC00-\uDD23\uDE80-\uDEA9\uDEAD-\uDF45\uDF51-\uDF81\uDF86-\uDFF6]"
        r"|\uD83A[\uDC00-\uDCCF\uDD00-\uDD43\uDD4B-\uDFFF]"
        r"|\uD83B[\uDC00-\uDEBB])",
    )


CONTAINS_RTL: Pattern = make_contains_rtl()


def contains_rtl(s: str) -> bool:
    """
    Returns true if `s` contains any Unicode character that is classified as "R" or "AL".
    """
    return CONTAINS_RTL.search(s) is not None


# Define a regular expression pattern for unusual line terminators
UNUSUAL_LINE_TERMINATORS: Pattern = re.compile(r"[\u2028\u2029]")


def contains_unusual_line_terminators(s: str) -> bool:
    """
    Returns true if `s` contains unusual line terminators, like LS or PS.
    """
    return UNUSUAL_LINE_TERMINATORS.search(s) is not None


def create_line_starts(r: list[int], s: str) -> tuple[list[int], int, int, int, bool]:
    """
    计算字符串中每一行的起始索引，并统计各种换行符的数量。

    Args:
        r: 用于存储行起始索引的列表，初始应为空或重置。
        s: 要分析的字符串。

    Returns:
        tuple[list[int], int, int, int, bool]: 包含行起始索引列表、CR数量、LF数量、CRLF数量和是否为基础ASCII的元组。
    """
    r.clear()
    r.append(0)  # 第一行总是从索引0开始
    cr = lf = crlf = 0  # 分别计数\r, \n, \r\n
    is_basic_ascii = True

    i = 0
    length = len(s)
    while i < length:
        i_chr = ord(s[i])
        if i_chr == 13:  # \r
            if i + 1 < length and ord(s[i + 1]) == 10:  # \r\n
                crlf += 1
                r.append(i + 2)
                i += 1  # 跳过\n
            else:
                cr += 1
                r.append(i + 1)
        elif i_chr == 10:  # \n
            lf += 1
            r.append(i + 1)
        else:
            # 检查是否为基础ASCII
            if is_basic_ascii:
                if i_chr != 9 and (i_chr < 32 or i_chr > 126):
                    is_basic_ascii = False
        i += 1

    # 返回构造的结果
    return r, cr, lf, crlf, is_basic_ascii


def is_pure_basic_ascii(lines: list[str]) -> bool:
    """
    检查给定的字符串列表中的所有文本是否全部由基础ASCII字符组成。

    参数:
    lines: 字符串列表。

    返回:
    如果列表中的所有字符串仅包含基础ASCII字符，则返回True；否则返回False。
    """
    for line in lines:
        for char in line:
            if ord(char) < 32 or ord(char) > 126:
                if ord(char) != 9:  # 允许制表符
                    return False
    return True


def count_eol(text: str) -> tuple[int, int, int, int]:
    """
    Counts the number of end-of-line characters and types in the text, and returns the length of the first line and the
    start of the last line.

    统计文本中的行结束符数量和类型，并返回首行长度及最后一行的起始索引。

    Args:
        text: A string representing the input text. | 输入文本的字符串。

    Returns:
        A tuple containing four integers:
        返回一个包含四个整数的元组：

        1. eol_count: The number of end-of-line characters or sequences in the text. | 文本中行结束符的数量。
        2. first_line_length: The length of the first line up to but not including the first end-of-line character. |
            首行的长度，不包括第一个行结束符。
        3. last_line_start_to_end_length: The length of the text from the start of the last line to the end of the text.
            | 从最后一行开始到文本结束的长度。
        4. eol: An integer representing the type of end-of-line character(s) found in the text. Possible values are:
            | 表示在文本中找到的行结束符类型的整数。可能的值包括：
            - 0: No end-of-line character found. | 没有找到行结束符。
            - 1: Only the line feed character ('\n') found. | 仅找到换行符（'\n'）。
            - 2: Only the carriage return and line feed sequence ('\r\n') found. | 仅找到回车和换行符序列（'\r\n'）。
            - 3: Other combinations or invalid end-of-line character(s) found. | 找到其他组合或无效的行结束符。
    """
    eol_count = 0
    first_line_length = 0
    last_line_start = 0
    eol = 0  # StringEOL.Unknown

    i = 0
    length = len(text)
    while i < length:
        i_chr = ord(text[i])
        if i_chr == 13:  # CharCode.CarriageReturn
            if eol_count == 0:
                first_line_length = i
            eol_count += 1
            if i + 1 < length and ord(text[i + 1]) == 10:  # CharCode.LineFeed
                eol |= 2  # StringEOL.CRLF
                i += 1  # skip \n
            else:
                eol |= 3  # StringEOL.Invalid
            last_line_start = i + 1
        elif i_chr == 10:  # CharCode.LineFeed
            eol |= 1  # StringEOL.LF
            if eol_count == 0:
                first_line_length = i
            eol_count += 1
            last_line_start = i + 1
        i += 1

    if eol_count == 0:
        first_line_length = length

    return eol_count, first_line_length, length - last_line_start, eol
