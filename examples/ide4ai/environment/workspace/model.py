# filename: model.py
# @Time    : 2024/4/18 12:25
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import copy
import math
import re
import uuid
from collections.abc import Callable, Iterator, Sequence
from pathlib import Path
from typing import Literal, cast

from cachetools import LRUCache
from jinja2 import Template
from pydantic import AnyUrl, FilePath
from typing_extensions import TypedDict

from ide4ai.environment.workspace.common.dispose import (
    DisposableProtocol,
)
from ide4ai.environment.workspace.schema import (
    Cursor,
    CursorStateComputer,
    EndOfLinePreference,
    EndOfLineSequence,
    IdentifiedSingleEditOperation,
    ModelContentChangedEvent,
    ModelProtocol,
    Position,
    Range,
    SearchResult,
    SingleEditOperation,
    TextChange,
    TextEdit,
    TextModelResolvedOptions,
    WordAtPosition,
)
from ide4ai.environment.workspace.utils import (
    LARGE_FILE_HEAP_OPERATION_THRESHOLD,
    contains_rtl,
    contains_unusual_line_terminators,
    count_eol,
    detect_newline_type,
    first_non_whitespace_index,
    is_pure_basic_ascii,
    last_non_whitespace_index,
    normalize_indentation,
    read_file_with_bom_handling,
)
from ide4ai.schema import LanguageId

MODEL_ID = 0

DEFAULT_CREATION_OPTIONS = {
    "tab_size": 4,
    "indent_size": 4,
    "insert_spaces": True,
    "default_eol": 1,  # Model.DefaultEndOfLine.LF
    "trim_auto_whitespace": True,
}

# DEFAULT_VIEW_TEMPLATE = (
#     "以下是{uri}的文档内容:\n```{language_id}\n{content_value}\n```\n"
#     "请注意为了方便编辑，我们使用格式化的文本将光标所在位置做了标注。光标的格式为: "
#     '">{{key}}|{{line_number}}:{{column_number}}<"\n'
#     "你计算位置时，需要忽略光标结构对字符串长度带来的影响，比如：\n"
#     "```python\n"
#     ">primary|1:1<print('hello world!')\n"
#     "```\n"
#     "在这个例子中，第一行的字符串的长度是22，但是行首光标的位置是1:1，因为光标的位置是在字符串的左边界，同时行末位置应该是1:22，
#     EOL并不算入行内字符。\n"
#     "为了方便你阅读行号，在每一行的开始，插入了一个宽度为5的字符串记录当前行号，与代码正文使用 '|' 分隔，
#     计算光标位置的时候不要计算这一部分。光标定位坐标采用 1-based，注意均是从1开始计数\n"
# )

DEFAULT_VIEW_TEMPLATE = (
    "以下是{uri}的文档内容:\n```{language_id}\n{original_content}\n```\n"
    "请注意为了方便编辑，我们使用格式化的文本将光标所在位置做了标注，同时每一行添加了行号备注。光标的格式为: "
    '">{{key}}|{{line_number}}:{{column_number}}<"\n'
    "你计算位置时，需要忽略光标结构对字符串长度带来的影响，比如：\n"
    "```{language_id}\n"
    ">primary|1:1<print('hello world!')\n"
    "```\n"
    "在这个例子中，第一行的字符串的长度是22，不会受光标结构影响（即光标不占用实际宽度）。行首光标的位置是1:1，同时行末位置应该是1:22，EOL并不算入行内字符。\n"
    "为了方便你阅读行号，在每一行的开始，插入了一个固定宽度的字符串记录当前行号，与代码正文使用 '|' 分隔，"
    "计算光标位置的时候不要计算这一部分。光标定位坐标采用 1-based，注意均是从1开始计数\n"
    "添加了光标与行号信息后的文本如下:\n```{language_id}\n{content_value}\n```\n"
)


DEFAULT_SIMPLE_VIEW_TEMPLATE = (
    "以下是{uri}的文档内容:\n```{language_id}\n{content_value}\n```\n"
    "为了方便你阅读行号，在每一行的开始，插入了一个固定宽度的字符串记录当前行号，与代码正文使用 '|' 分隔，行列号均为 1-based\n"
)


class OperationDict(TypedDict):
    sort_index: int
    identifier: str
    range: Range
    range_offset: int
    range_length: int
    text: str
    eol_count: int
    first_line_length: int
    last_line_length: int
    force_move_markers: bool
    is_auto_whitespace_edit: bool


class TextModel(ModelProtocol):
    """
    当前版本的TextModel不需要支持Undo-Redo。因为是给大模型使用，所以多数操作其可以正常生成指令，它区别于人类，思考一个指令和触发一个
    Undo-Redo的成本几乎是一致的。

    TODO: 需要对SurrogatePairs的情况进行兼容。需要尽快支持Undo-Redo操作。需要尽快支持搜索操作。
    """

    def __init__(
        self,
        language_id: LanguageId,
        creation_options: dict | None = None,
        uri: AnyUrl | None = None,
        view_template: str = DEFAULT_VIEW_TEMPLATE,
        simple_view_template: str = DEFAULT_SIMPLE_VIEW_TEMPLATE,
        auto_save_during_dispose: bool = True,
    ) -> None:
        self._is_disposed: bool = False  # 设置是否已释放，默认为False
        self._is_disposing: bool = False
        self._auto_save_during_dispose: bool = auto_save_during_dispose
        self.__version_id: int = 1
        self.__language_id: LanguageId = language_id  # The language identifier of the text model.
        self.view_template = view_template
        self.simple_view_template = simple_view_template
        options: dict = copy.deepcopy(DEFAULT_CREATION_OPTIONS)  # The default creation options of the text model.
        if creation_options:
            options.update(creation_options)
        self.creation_options: TextModelResolvedOptions = TextModelResolvedOptions(**options)
        global MODEL_ID
        MODEL_ID += 1
        if not uri:
            uri = AnyUrl(f"inmemory://model/{MODEL_ID}")
        self.__uri: AnyUrl = uri  # The URI of the text model.
        self.__m_id: str = f"$model{MODEL_ID}"  # The ID of the text model.
        self._content: list[str] = []  # The content lines of the text model.
        if not (str(uri).startswith("inmemory://") or str(uri).startswith("file://")):
            raise ValueError("Only 'inmemory://' or 'file://' URIs are supported now.")
        elif str(uri).startswith("inmemory://"):
            self.__eol: EndOfLineSequence = EndOfLineSequence.LF
            self.__bom = ""
            self._content = [""]
        else:
            path = uri.path
            if not path:
                raise ValueError("The path of the URI is empty.")
            self.__eol = detect_newline_type(Path(path))
            self.__bom, self._content = read_file_with_bom_handling(path)
        self._is_basic_ascii = is_pure_basic_ascii(self._content)
        if not self._is_basic_ascii:
            self._might_contain_RTL = any(contains_rtl(text) for text in self._content)
            self._might_contain_unusual_line_terminators = any(
                contains_unusual_line_terminators(text) for text in self._content
            )
        else:
            self._might_contain_unusual_line_terminators = False
            self._might_contain_RTL = False
        self.cursors: LRUCache = LRUCache(maxsize=10)  # The cursors of the text model. | 文本模型的光标，至多缓存10个。
        self.cursors["primary"] = Cursor(key="primary", position=Position(1, 1))

    @property
    def uri(self) -> AnyUrl:
        """
        The associated URI for this document. Most documents have the file-scheme, indicating that they represent files
        on disk. However, some documents may have other schemes indicating that they are not available on disk.

        Returns:

        """
        return self.__uri

    @property
    def m_id(self) -> str:
        """
        The ID of the model.

        Returns:
            str: The ID of the model.
        """
        return self.__m_id

    @property
    def eol(self) -> EndOfLineSequence:
        return self.__eol

    @property
    def language_id(self) -> LanguageId:
        return self.__language_id

    def get_language_id(self) -> str:
        """
        Get the language id of the model.

        Returns:
            str: The language id of the model.
        """
        return self.__language_id.value

    def get_options(self) -> TextModelResolvedOptions:
        return self.creation_options

    def get_version_id(self) -> int:
        self._assert_not_disposed()
        return self.__version_id

    def _increase_version_id(self) -> None:
        """
        Increments the version ID by 1.

        This method increases the value of the version ID attribute by 1.
        The version ID is a private attribute denoted by the double underscore prefix.
        """
        self.__version_id += 1

    def _overwrite_version_id(self, version_id: int) -> None:
        """
        Overwrites the version ID with the given value.

        Args:
            version_id (int): The new version ID to set.
        """
        self.__version_id = version_id

    def get_alternative_version_id(self) -> int:
        """
        当前版本不支持Undo-Redo，故而不需要实现此方法。

        Returns:
            int: The alternative version ID.
        """
        return self.__version_id

    def set_value(self, new_value: str | Iterator[str]) -> None:
        """
        Replace the entire text buffer value contained in this model.

        Args:
            new_value (str | Iterator[str]): The new value to set.
        """
        self._assert_not_disposed()
        if isinstance(new_value, str):
            new_value_bytes = new_value.encode("utf-8")
            if len(new_value_bytes) > LARGE_FILE_HEAP_OPERATION_THRESHOLD:
                raise ValueError("File content exceeds the memory usage threshold of 256K characters.")
            self._content = new_value.splitlines()
        else:
            total_bytes = 0
            self._content = []
            for line in new_value:
                line_bytes = line.encode("utf-8")
                total_bytes += len(line_bytes)
                if total_bytes > LARGE_FILE_HEAP_OPERATION_THRESHOLD:
                    raise ValueError("File content exceeds the memory usage threshold of 256K characters.")
                self._content.append(line)
        if not self._content:
            # 因为存在空内容的情况，空内容也需要有一个无任何内容的空行，方可输入，否则光标无处定位
            self._content = [""]
        self._increase_version_id()

    def save(self, path: FilePath | None = None) -> None:
        """
        将 self._content 中的内容写入到文件

        Returns:

        """
        self._assert_not_disposed()
        if not path:
            if self.uri.scheme == "file" and self.uri.path:
                path = Path(self.uri.path)
            else:
                raise ValueError("You should provide a path to save the file.")
        # 写入
        with open(str(path), "w", encoding="utf-8") as f:
            f.write(self.get_value())

    def disable_auto_save(self) -> None:
        self._auto_save_during_dispose = False

    def enable_auto_save(self) -> None:
        self._auto_save_during_dispose = True

    def dispose(self) -> None:
        if self._is_disposed:
            return  # 避免重复释放。dispose可能会注册到多个位置，比如atexit，因为如果手动释放过后，再遇到自动触发的释放，则直接return
        self._is_disposing = True
        if self._auto_save_during_dispose:
            self.save()
        self._content.clear()
        self.cursors.clear()
        self.__version_id = 1
        self._is_disposed = True
        self._is_disposing = False

    def _assert_not_disposed(self) -> None:
        if self._is_disposed:
            raise ValueError("Text model is already disposed.")

    def get_value(
        self,
        eol: EndOfLinePreference | None = EndOfLinePreference.TEXT_DEFINED,
        preserve_bom: bool | None = False,
    ) -> str:
        """
        Get the entire text buffer value contained in this model. | 获取此模型中包含的整个文本缓冲区值。

        Args:
            eol (Optional[EndOfLinePreference]): The end of line character preference. Defaults to
                EndOfLinePreference.TextDefined. | 行尾字符首选项。默认为EndOfLinePreference.TextDefined。
            preserve_bom (Optional[bool]): Preserve a BOM character if it was detected when the model was constructed.
            |  如果在构造模型时检测到BOM字符，则保留该字符。

        Returns:
            str: The entire text buffer value.

        BOM是字节顺序标记(Byte Order Mark)的缩写。它是用于指示一个文本流的字节顺序的Unicode字符。具体来说，BOM是一个位于文本开头的特殊字符，
        该字符用于表明文件是用UTF-8、UTF-16还是UTF-32编码，并且表明该文件是用大端字节序还是小端字节序编码。在UTF-16和UTF-32中，BOM是必需的，
        因为这两种编码都包含非ASCII字符，这些字符的字节顺序不同。但是，在UTF-8中，BOM是可选的，因为UTF-8字符的字节顺序是固定的。然而，
        不管编码方式如何，有些文本处理软件在读取文本文件时，都会检查文件的开头是否包含BOM。如果有BOM，则按照BOM的指示来解码文件。如果没有BOM，
        那么这个软件必须根据其他信息（如默认设置或用户输入）来判断文件的编码方式。因此，即使是UTF-8文本，有时候也会包含一个BOM，以方便那些需要BOM的软件处理。
        """
        if eol is None:
            eol = EndOfLinePreference.TEXT_DEFINED
        self._assert_not_disposed()
        full_range = self.get_full_model_range()
        full_value = self.get_value_in_range(full_range, eol)
        if preserve_bom:
            return self.__bom + full_value
        return full_value

    def create_snapshot(self, preserve_bom: bool | None = None) -> Iterator[str]:
        raise NotImplementedError

    def get_value_length(
        self,
        eol: EndOfLinePreference | None = None,
        preserve_bom: bool | None = None,
    ) -> int:
        """
        Get the length of the text stored in this model.

        Args:
            eol (Optional[EndOfLinePreference]): The end of line character preference. Defaults to
                EndOfLinePreference.TextDefined.
            preserve_bom (Optional[bool]): Preserve a BOM character if it was detected when the model was constructed.

        Returns:
            int: The length of the text stored in this model.
        """
        self._assert_not_disposed()
        full_model_range = self.get_full_model_range()
        full_model_value_len = self.get_value_length_in_range(full_model_range, eol)
        if preserve_bom:
            return len(self.__bom) + full_model_value_len
        return full_model_value_len

    def get_value_in_range(
        self,
        t_range: Range,
        eol: EndOfLinePreference | None = EndOfLinePreference.TEXT_DEFINED,
    ) -> str:
        """
        Get the text stored in this model by the given range.

        Args:
            t_range (Range): The range of the text.
            eol (EndOfLinePreference): The end of line character preference. This will only be used for multiline
                ranges. Defaults to EndOfLinePreference.TextDefined.

        Returns:
            str: The text stored in this model.
        """
        if eol is None:
            eol = EndOfLinePreference.TEXT_DEFINED
        if t_range.is_empty():
            return ""
        content_eol = self.get_eol()
        start_line = t_range.start_position.line
        end_line = t_range.end_position.line
        start_column = t_range.start_position.character
        end_column = t_range.end_position.character
        if start_line == end_line:
            value: str = self._content[start_line - 1][start_column - 1 : end_column - 1]
        else:
            value = self._content[start_line - 1][start_column - 1 :]
            for i in range(start_line + 1, end_line):
                value += content_eol + self._content[i - 1]
            value += content_eol + self._content[end_line - 1][: end_column - 1]
        line_ending = self.get_eol_by_preference(eol)
        if line_ending != content_eol:
            value = value.replace(content_eol, line_ending)
        return value

    def get_value_length_in_range(self, t_range: Range, eol: EndOfLinePreference | None) -> int:
        """
        Get the length of the text stored in this model by the given range.

        Args:
            t_range (Range): The range of the text.
            eol (EndOfLinePreference): The end of line character preference. This will only be used for multiline
                ranges. Defaults to EndOfLinePreference.TextDefined.

        Returns:
            int: The length of the text stored in this model.
        """
        if t_range.is_empty():
            return 0
        if t_range.start_position.line == t_range.end_position.line:
            return t_range.end_position.character - t_range.start_position.character
        start_offset = self.get_offset_at(t_range.start_position)
        end_offset = self.get_offset_at(t_range.end_position)
        eol_offset_compensation = 0
        actual_eol = self.get_eol()
        desired_eol = self.get_eol_by_preference(eol)
        if desired_eol != actual_eol:
            delta = len(desired_eol) - len(actual_eol)
            eol_offset_compensation = delta * (t_range.end_position.line - t_range.start_position.line)
        return end_offset - start_offset + eol_offset_compensation

    def get_full_model_range(self) -> Range:
        """
        Get a range covering the entire model.

        Returns:
            Range: The full model range.
        """
        self._assert_not_disposed()
        line_count = self.get_line_count()
        return Range(
            start_position=Position(line=1, character=1),
            end_position=Position(line=line_count, character=self.get_line_max_column(line_count)),
        )

    def get_line_count(self) -> int:
        """
        Get the number of lines in the model.

        Returns:
            int: The number of lines in the model.
        """
        return len(self._content)

    def _is_valid_range(self, t_range: Range) -> bool:
        """
        If the t_range could pass validation

        Args:
            t_range (Range): The range to validate.

        Returns:
            bool: If the range is valid.
        """
        if not self._is_valid_position(t_range.start_position) or not self._is_valid_position(t_range.end_position):
            return False
        return True

    def validate_range(self, t_range: Range) -> Range:
        """
        Validate the range. | 验证范围合法性。

        Args:
            t_range (Range): The range to validate.

        Returns:
            Range: The validated range.
        """
        self._assert_not_disposed()
        if self._is_valid_range(t_range):
            return t_range
        start = self._validate_position(t_range.start_position)
        end = self._validate_position(t_range.end_position)
        return Range(start_position=start, end_position=end)

    def get_offset_at(self, position: Position) -> int:
        """
        Get the offset at the given position.

        Notes:
            请注意offset包括了换行符的宽度！！！

        Args:
            position (Position): The position to get the offset.

        Returns:
            int: The offset at the given position.
        """
        self._assert_not_disposed()
        position = self._validate_position(position)
        line_number = position.line
        offset = 0
        current_eol = self.get_eol()
        for line_num in range(line_number):
            if line_num == line_number - 1:
                offset += position.character - 1
            else:
                offset += len(self._content[line_num]) + len(current_eol)
        return offset

    def get_position_at(self, offset: int) -> Position:
        """
        Get the position at the given offset.

        Notes:
            请注意offset包括了换行符的宽度！！！

        Args:
            offset (int): The offset to get the position.

        Returns:
            Position: The position at the given offset.
        """
        self._assert_not_disposed()
        if offset <= 0:
            return Position(1, 1)
        line_count = self.get_line_count()
        if line_count == 0:
            return Position(1, 1)
        if offset >= self.get_value_length():
            return Position(line_count, self.get_line_max_column(line_count))
        line_number = 1
        current_eol = self.get_eol()
        while offset >= 0:
            line_length = len(self._content[line_number - 1]) + len(current_eol)
            if offset < line_length:
                return Position(line_number, offset + 1)
            offset -= line_length
            line_number += 1
        return Position(1, 1)

    def _is_valid_position(self, position: Position) -> bool:
        """
        判断Position是否合法

        Args:
            position (Position): The position to validate.

        Returns:
            bool: if pass validation
        """
        line_number = position.line
        column = position.character
        if not isinstance(line_number, int) or not isinstance(column, int):
            return False
        if line_number < 1 or column < 1:
            return False
        if line_number != int(line_number) or column != int(column):
            return False
        line_count = self.get_line_count()
        if line_number > line_count:
            return False
        if column == 1:
            return True
        max_column = self.get_line_max_column(line_number)
        if column > max_column:
            return False
        return True

    def _validate_position(self, position: Position) -> Position:
        """
        Validate the position. | 验证位置合法性。

        Notes:
            注意我们并未处理SurrogatePairs对应的情况，因为使用Python Open函数打开一个文件的时候，SurrogatePairs已经被处理过了。

        Args:
            position (Position): The position to validate.

        Returns:
            Position: The validated position.
        """
        _line_number = position.line
        _column = position.character
        line_number = max(
            1,
            int(_line_number) if isinstance(_line_number, (int, float)) and not math.isnan(_line_number) else 1,
        )
        column = max(
            1,
            int(_column) if isinstance(_column, (int, float)) and not math.isnan(_column) else 1,
        )

        lineCount = self.get_line_count()
        if line_number > lineCount:
            return Position(lineCount, self.get_line_max_column(lineCount))
        elif line_number < 1:
            return Position(1, 1)

        if column <= 1:
            return Position(line_number, 1)

        maxColumn = self.get_line_max_column(line_number)
        if column >= maxColumn:
            return Position(line_number, maxColumn)

        return Position(line_number, column)

    def validate_position(self, position: Position) -> Position:
        """
        Validate the position. | 验证位置合法性。

        Args:
            position (Position): The position to validate.

        Returns:
            Position: The validated position.
        """
        self._assert_not_disposed()
        if self._is_valid_position(position):
            return position
        else:
            return self._validate_position(position)

    def find_matches(
        self,
        search_string: str,
        search_scope: Range | list[Range] | None = None,
        is_regex: bool = False,
        match_case: bool = False,
        word_separator: str | None = None,
        capture_matches: bool = False,
        limit_result_count: int | None = None,
    ) -> list[SearchResult]:
        """
            Search for matches in the model based on the provided search string and parameters.

            根据提供的搜索字符串和参数在模型中搜索匹配项。

        Args:
            search_string: The string to search for. | 要搜索的字符串。
            search_scope: Optional. The range or list of ranges where the search should be performed. If not provided,
                the search will be performed in the full model range. | 可选。指定搜索应在其中进行的范围或范围列表。如果未提供，
                则在整个模型范围内进行搜索。
            is_regex: Optional. Specifies whether the search string should be treated as a regular expression. Default
                is False. | 可选。指定是否应将搜索字符串视为正则表达式。默认为 False。
            match_case: Optional. Specifies whether the search should be case-sensitive. Default is False. | 可选。指定搜
                索是否应区分大小写。默认为 False。
            word_separator: Optional. The separator used to define word boundaries in the search. If not provided, all
                characters are considered as part of a word. | 可选。用于定义搜索中单词边界的分隔符。如果未提供，
                则所有字符都视为单词的一部分。
            capture_matches: Optional. Specifies whether the matched ranges should be captured in the search results.
                Default is False. | 可选。指定是否应在搜索结果中捕获匹配的范围。默认为 False。
            limit_result_count: Optional. The maximum number of search results to return. If not provided, all matches
                will be returned. | 可选。返回的搜索结果的最大数量。如果未提供，将返回所有匹配项。

        Returns:
            A list of Range objects representing the matched ranges. | 表示匹配范围的 Range 对象列表。

        Raises:
            ValueError: If an invalid search scope is provided. | 如果提供了无效的搜索范围。
        """
        self._assert_not_disposed()
        # Handle search scope and set default to full model range if None
        search_ranges: list[Range]
        if search_scope is None:
            search_ranges = [self.get_full_model_range()]
        elif isinstance(search_scope, Range):
            search_ranges = [search_scope]
        elif isinstance(search_scope, list):
            search_ranges = search_scope
        else:
            raise ValueError("Invalid search scope.")  # pragma: no cover
        # Sort and merge intersecting ranges
        search_ranges.sort(key=lambda r: (r.start_position.line, r.start_position.character))
        unique_search_ranges = []
        current_range = search_ranges[0]

        for t_range in search_ranges[1:]:
            if current_range & t_range:
                current_range |= t_range
            else:
                unique_search_ranges.append(current_range)
                current_range = t_range
        unique_search_ranges.append(current_range)

        if not is_regex and "\n" not in search_string:
            # Handle non-regex, non-multiline case
            return self.find_matches_line_by_line(
                unique_search_ranges,
                search_string,
                match_case,
                capture_matches,
                limit_result_count,
            )

        # Prepare the search parameters
        from ide4ai.environment.workspace.model_search import (
            SearchParams,
            TextModelSearch,
        )

        search_params = SearchParams(
            search_string=search_string,
            is_regex=is_regex,
            match_case=match_case,
            word_separators=word_separator,
        )
        # Delegate the actual matching to the static method in TextModelSearch

        results: list[SearchResult] = []
        for search_range in unique_search_ranges:
            results += TextModelSearch.find_matches(
                self,
                search_params,
                search_range,
                capture_matches,
                limit_result_count,
            )
        return results

    def find_matches_line_by_line(
        self,
        search_scope: list[Range],
        search_string: str,
        match_case: bool,
        capture_matches: bool,
        limit_result_count: int | None,
    ) -> list[SearchResult]:
        """
        Args:
            search_scope (list[Range]): The list of ranges where the search should be  performed. | 搜索应该执行的范围列表。
            search_string (str): The string to search for. | 要搜索的字符串。
            match_case (bool): 表示搜索是否区分大小写的布尔值。
                A boolean indicating whether the search should be case-sensitive or not.
            capture_matches (bool): 表示在结果中是否应该捕获匹配的子字符串的布尔值。
                A boolean indicating whether the matched substrings should be captured in the results.
            limit_result_count (int, optional): 表示要返回的最大匹配数的整数。如果提供，搜索将在找到这些数量的匹配后停止。
                An optional integer indicating the maximum number of matches to return. If provided,
                the search will stop after finding this number of matches.

        Returns:
            list[SearchResult]: 每个找到的匹配都表示为一个字典。每个字典包含行号，开始索引，结束索引，以及匹配的文本（如果capture_matches
                为True）。如果提供了limit_result_count并且匹配数量超过了它，函数将返回到那个点为止找到的匹配。
                A list of dictionaries representing each match found. Each dictionary contains the line number, start
                index, end index, and the matched text (if capture_matches is True). If limit_result_count is provided
                and the number of matches exceeds it, the function will return the matches found until that point.

        """
        # 初始化结果列表和总匹配数 | Initialize results list and total match count
        results: list[SearchResult] = []
        total_matches = 0
        # 如果匹配不区分大小写，将搜索字符串转化为小写 | Convert the search string to lowercase if the match is case-insensitive
        search_string = search_string if match_case else search_string.lower()
        # 枚举内容的每一行
        # Enumerate over each line in the content
        for r in search_scope:
            range_scope = self.validate_range(r)
            find_scope_content = self.get_value_in_range(range_scope)
            contents = find_scope_content.splitlines(keepends=False)
            for line_number, line in enumerate(contents, start=range_scope.start_position.line):
                # 如果匹配不区分大小写，将该行转化为小写 | Convert the line to lowercase if the match is case-insensitive
                if not match_case:
                    line = line.lower()

                # 初始化开始位置
                # Initialize start position
                start = 0
                while True:
                    # 查找搜索字符串的位置 | Find the position of search string
                    start = line.find(search_string, start)
                    if start == -1:
                        break
                    # 确定结束位置 | Determine end position
                    end = start + len(search_string)
                    # 获取匹配的文本 | Get the matched text
                    match_text = line[start:end]
                    # 添加搜索结果到结果列表中 | Append search result to results list
                    results.append(
                        SearchResult(
                            range=Range(
                                start_position=Position(line=line_number, character=start + 1),
                                end_position=Position(line=line_number, character=end + 1),
                            ),
                            match=match_text if capture_matches else None,
                        ),
                    )
                    # 移动开始位置到刚刚过去的最后一个匹配 | Move start to just past the last match
                    start += len(search_string)
                    # 增加总匹配数
                    # Increase total match count
                    total_matches += 1
                    # 如果达到限制的结果数量，返回结果 | If limit on the result count is reached, return results
                    if limit_result_count and total_matches >= limit_result_count:
                        return results
        return results

    def find_next_match(
        self,
        search_string: str,
        search_start: Position,
        search_mode: bool | Range | list[Range],
        is_regex: bool,
        match_case: bool,
        word_separator: str | None = None,
        capture_matches: bool = False,
    ) -> Range | None:
        raise NotImplementedError

    def find_previous_match(
        self,
        search_string: str,
        search_start: Position,
        search_mode: bool | Range | list[Range],
        is_regex: bool,
        match_case: bool,
        word_separator: str | None = None,
        capture_matches: bool = False,
    ) -> Range | None:
        raise NotImplementedError

    def get_word_at_position(self, position: Position) -> WordAtPosition | None:
        """
        Get the word under or besides position.

        Args:
            position (Position): The position to get the word.

        Returns:
            Optional[WordAtPosition]: The word at the given position, or None if no word is found.
        """
        raise NotImplementedError

    def get_word_until_position(self, position: Position) -> WordAtPosition:
        """
        Get the word until position.

        Args:
            position (Position): The position to get the word.

        Returns:
            WordAtPosition: The word until the given position.
        """
        raise NotImplementedError

    def update_options(self, options: TextModelResolvedOptions) -> None:
        """
        Update the options of the model.

        Args:
            options (TextModelResolvedOptions): The options to update.
        """
        self._assert_not_disposed()
        self.creation_options = options

    def push_stack_element(self) -> None:
        raise NotImplementedError

    def pop_stack_element(self) -> None:
        raise NotImplementedError

    def push_eol(self, eol: EndOfLineSequence) -> None:
        raise NotImplementedError

    def push_edit_operations(
        self,
        before_cursor_state: list[Range] | None,
        edit_operations: list[SingleEditOperation],
        cursorStateComputer: CursorStateComputer,
    ) -> list[Range] | None:
        raise NotImplementedError

    def set_eol(self, eol: EndOfLineSequence) -> None:
        self._assert_not_disposed()
        new_rol = "\r\n" if eol == EndOfLineSequence.CRLF else "\n"
        if new_rol == self.get_eol():
            return
        self._increase_version_id()
        self.__eol = eol

    def get_eol(self) -> Literal["\r\n", "\n"]:
        """
        Get the end of line character.

        Returns:
            Literal["\r\n", "\n"]
        """
        return "\n" if self.__eol == EndOfLineSequence.LF else "\r\n"

    def get_eol_by_preference(self, eol: EndOfLinePreference | None) -> Literal["\r\n", "\n"]:
        """
        通过指定的EndOfLinePreference获取行尾字符。

        Args:
            eol (Optional[EndOfLinePreference]): The end of line preference.

        Returns:
            Literal["\r\n", "\n"]: The end of line character.
        """
        match eol:
            case EndOfLinePreference.LF:
                return "\n"
            case EndOfLinePreference.CRLF:
                return "\r\n"
            case EndOfLinePreference.TEXT_DEFINED:
                return "\n" if self.__eol == EndOfLineSequence.LF else "\r\n"
            case _:
                return "\n"

    def get_end_of_line_sequence(self) -> EndOfLineSequence:
        """
        Get the end of line sequence in the model.

        Returns:
            EndOfLineSequence: The end of line sequence in the model.
        """
        self._assert_not_disposed()
        return EndOfLineSequence.LF if self.get_eol() == "\n" else EndOfLineSequence.CRLF

    def get_line_min_column(self, line_num: int) -> int:
        """
        Get the minimum column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The minimum column of the line.
        """
        self._assert_not_disposed()
        if line_num < 1 or line_num > len(self._content):
            raise ValueError("Invalid line number.")
        return 1

    def get_line_max_column(self, line_num: int) -> int:
        """
        Get the maximum column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The maximum column of the line.
        """
        self._assert_not_disposed()
        if line_num < 1 or line_num > len(self._content):
            raise ValueError("Invalid line number.")
        return len(self._content[line_num - 1]) + 1

    def get_line_content(self, line_number: int) -> str:
        """
        Get the content of a line.

        Args:
            line_number (int): The line number.

        Returns:
            str: The content of the line.
        """
        self._assert_not_disposed()
        if line_number < 1 or line_number > len(self._content):
            raise ValueError("Invalid line number.")
        return self._content[line_number - 1]

    def get_line_length(self, line_number: int) -> int:
        """
        Get the length of a line.

        Args:
            line_number (int): The line number.

        Returns:
            int: The length of the line.
        """
        self._assert_not_disposed()
        return len(self.get_line_content(line_number))

    def get_lines_content(self) -> list[str]:
        """
        Get the content of all lines.

        Returns:
            list[str]: The content of all lines.
        """
        self._assert_not_disposed()
        return self._content

    def get_line_first_non_whitespace_column(self, line_num: int) -> int:
        """
        Get the first non-whitespace column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The first non-whitespace column of the line.
        """
        self._assert_not_disposed()
        if line_num < 1 or line_num > len(self._content):
            raise ValueError("Invalid line number.")
        column = first_non_whitespace_index(self.get_line_content(line_num))
        return column + 1

    def get_line_last_non_whitespace_column(self, line_num: int) -> int:
        """
        Get the last non-whitespace column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The last non-whitespace column of the line.
        """
        self._assert_not_disposed()
        if line_num < 1 or line_num > len(self._content):
            raise ValueError("Invalid line number.")
        content = self.get_line_content(line_num)
        res = last_non_whitespace_index(content)
        return res + 2 if res != -1 else 0

    def modify_position(self, position: Position, offset: int) -> Position:
        """
        Modify the given position by the given offset.

        Advances the given position by the given offset (negative offsets are also accepted) and returns it as a new
        valid position. If the offset and position are such that their combination goes beyond the beginning or end of
        the model, throws an exception. If the offset is such that the new position would be in the middle of a
        multi-byte line terminator(such as \r\n), throws an exception.

        Args:
            position (Position): The position to modify.
            offset (int): The offset to modify the position.

        Returns:
            Position: The modified position.
        """
        self._assert_not_disposed()
        candidate = self.get_offset_at(position) + offset
        return self.get_position_at(min(self.get_value_length(), max(0, candidate)))

    def normalize_indentation(self, text: str) -> str:
        """
        Normalize a string containing whitespace according to indentation rules (converts to spaces or to tabs).

        Args:
            text (str): The text to normalize.

        Returns:
            str: The normalized text.
        """
        self._assert_not_disposed()
        return normalize_indentation(text, self.creation_options.indent_size, self.creation_options.insert_spaces)

    def detect_indentation(self, default_insert_spaces: bool, default_tab_size: int) -> None:
        raise NotImplementedError

    def _validate_edit_operation(
        self,
        raw_operation: IdentifiedSingleEditOperation | SingleEditOperation,
    ) -> IdentifiedSingleEditOperation:
        """

        Args:
            raw_operation:

        Returns:

        """
        if isinstance(raw_operation, IdentifiedSingleEditOperation):
            return raw_operation

        return IdentifiedSingleEditOperation(
            identifier=getattr(raw_operation, "identifier", uuid.uuid4().hex),
            range=self.validate_range(raw_operation.range),
            text=raw_operation.text,
            force_move_markers=raw_operation.force_move_markers or False,
            is_auto_whitespace_edit=getattr(raw_operation, "is_auto_whitespace_edit", False),
            is_tracked=getattr(raw_operation, "_is_tracked", False),
        )

    def _validate_edit_operations(
        self,
        raw_operations: Sequence[IdentifiedSingleEditOperation | SingleEditOperation],
    ) -> list[IdentifiedSingleEditOperation]:
        """
        Validate a list of edit operations, returning a list of `ValidAnnotatedEditOperation` instances.
        """
        result = []
        for operation in raw_operations:
            result.append(self._validate_edit_operation(operation))
        return result

    def apply_edits(
        self,
        operations: list[SingleEditOperation],
        compute_undo_edits: bool | None = None,
    ) -> list[TextEdit] | None:
        """
        Edit the model without adding the edits to the undo stack. This can have dire consequences on the undo stack!
        See @push_edit_operations for the preferred way.

        Args:
            operations (list[SingleEditOperation]): The operations to apply.
            compute_undo_edits (Optional[bool]): Whether to compute undo edits.

        Returns:
            Optional[list[TextEdit]]: The text edits. If desired, the inverse edit operations, that, when applied,
                will bring the model back to the previous state.
        """
        id_operations = self._validate_edit_operations(operations)
        return self.__apply_edits(id_operations, compute_undo_edits)

    def __apply_edits(
        self,
        raw_operations: list[IdentifiedSingleEditOperation],
        compute_undo_edits: bool | None = None,
    ) -> list[TextEdit] | None:
        """
        Apply the edits to the model.

        Args:
            raw_operations (list[IdentifiedSingleEditOperation]): The operations to apply.
            compute_undo_edits (Optional[bool]): Whether to compute undo edits.

        Returns:
            Optional[list[TextEdit]]: The text edits.
        """
        might_contain_RTL = self._might_contain_RTL
        might_contain_unusual_line_terminators = self._might_contain_unusual_line_terminators
        might_contain_non_basic_ascii = not self._is_basic_ascii
        # 开始准备操作数据
        operations: list[OperationDict] = []
        for index, op in enumerate(raw_operations):
            validate_range = op.range  # 通过在apply_edits中的调用保证了op中的range均是经过验证的range
            if op.text:  # 如果当前变更有指定内容。如果不指定内容则等同于删除
                text_might_contain_non_basic_ascii = True
                if not might_contain_non_basic_ascii:
                    text_might_contain_non_basic_ascii = not is_pure_basic_ascii([op.text])
                    might_contain_non_basic_ascii = text_might_contain_non_basic_ascii
                if not might_contain_RTL and text_might_contain_non_basic_ascii:
                    might_contain_RTL = contains_rtl(op.text)
                if not might_contain_unusual_line_terminators and text_might_contain_non_basic_ascii:
                    might_contain_unusual_line_terminators = contains_unusual_line_terminators(op.text)
            valid_text: str = ""
            eol_count: int = 0
            first_line_length: int = 0
            last_line_length: int = 0
            if op.text:
                eol_count, first_line_length, last_line_length, str_eol = count_eol(op.text)
                current_eol = self.get_eol()
                expected_str_eol = 2 if current_eol == "\r\n" else 1
                if str_eol == 0 or str_eol == expected_str_eol:
                    valid_text = op.text
                else:
                    valid_text = re.sub(r"\r\n|\r|\n", current_eol, op.text)
            operations.append(
                {
                    "sort_index": index,
                    "identifier": op.identifier,
                    "range": validate_range,
                    "range_offset": self.get_offset_at(validate_range.start_position),
                    "range_length": self.get_value_length_in_range(validate_range, None),
                    "text": valid_text,
                    "eol_count": eol_count,
                    "first_line_length": first_line_length,
                    "last_line_length": last_line_length,
                    "force_move_markers": op.force_move_markers or False,
                    "is_auto_whitespace_edit": op.is_auto_whitespace_edit,
                },
            )
        operations.sort(
            key=lambda x: x["range"],
            reverse=True,
        )  # Python自带的sorted方法与 .sort() 均是稳定排序，即相等的元素排序前后的顺序不变
        has_touching_ranges: bool = False  # 是否有相邻的Range操作恰好有相连但无交集
        count = len(operations) - 1
        for i in range(count):
            # 因为经过一轮倒序，所以大文档中越向后面的Range会在此数据越靠前。判断overlapping的时候，需要判断前一个Range的start_position
            # 是否小于后一个Range的end_position
            range_start = operations[i]["range"].start_position
            next_range_end = operations[i + 1]["range"].end_position
            if range_start <= next_range_end:
                if range_start < next_range_end:
                    # overlapping ranges
                    raise ValueError("Overlapping ranges are not allowed!")
                has_touching_ranges = True
        reverse_ranges = (
            self._get_inverse_edit_ranges(operations)
            if compute_undo_edits or self.creation_options.trim_auto_whitespace
            else []
        )
        new_trim_auto_whitespace_candidates = []
        if self.creation_options.trim_auto_whitespace:
            for index, op_dict in enumerate(operations):
                reverse_range = reverse_ranges[index]
                if op_dict["is_auto_whitespace_edit"] and op_dict["range"].is_empty():
                    # Record already the future line numbers that might be auto whitespace removal candidates on next
                    # edit 提前记录下一次编辑可能会自动移除空白的行号候选项。
                    for line_number in range(
                        reverse_range.start_position.line,
                        reverse_range.end_position.line + 1,
                    ):
                        current_line_content = ""
                        if line_number == reverse_range.start_position.line:
                            current_line_content = self.get_line_content(op_dict["range"].start_position.line)
                            if first_non_whitespace_index(current_line_content) != -1:
                                continue
                        new_trim_auto_whitespace_candidates.append(
                            {
                                "line_number": line_number,
                                "old_content": current_line_content,
                            },
                        )
        reverse_operations: list[dict] = []
        if compute_undo_edits:
            reverse_range_delta_offset = 0
            reverse_operations = []
            for i, op_dict in enumerate(operations):
                reverse_range = reverse_ranges[i]
                buffer_text = self.get_value_in_range(op_dict["range"])
                reverse_range_offset = op_dict["range_offset"] + reverse_range_delta_offset
                reverse_range_delta_offset += len(op_dict["text"]) - len(buffer_text)
                reverse_operations.append(
                    {
                        "sort_index": op_dict["sort_index"],
                        "identifier": op_dict["identifier"],
                        "range": reverse_range,
                        "text": buffer_text,
                        "text_change": TextChange(
                            old_position_offset=op_dict["range_offset"],
                            old_text=buffer_text,
                            new_position_offset=reverse_range_offset,
                            new_text=op_dict["text"],
                        ),
                    },
                )

            # Can only sort reverse operations when the order is not significant
            if not has_touching_ranges:
                reverse_operations.sort(key=lambda x: x["sort_index"])
        self._might_contain_RTL = might_contain_RTL
        self._might_contain_unusual_line_terminators = might_contain_unusual_line_terminators
        self._is_basic_ascii = not might_contain_non_basic_ascii
        # 结束准备操作数据
        # 开始执行操作
        self._do_apply_edits(operations)
        self._increase_version_id()
        # 结束执行操作
        trim_auto_whitespace_line_numbers = None
        if len(new_trim_auto_whitespace_candidates) > 0 and self.creation_options.trim_auto_whitespace:
            # 对候选行号按降序排序
            new_trim_auto_whitespace_candidates.sort(key=lambda x: cast(int, x["line_number"]), reverse=True)
            trim_auto_whitespace_line_numbers = []
            previous_line_number = None
            for candidate in new_trim_auto_whitespace_candidates:
                line_number = cast(int, candidate["line_number"])
                if line_number == previous_line_number:
                    # 不添加重复的行号
                    continue
                previous_line_number = line_number
                prev_content = candidate["old_content"]
                line_content = self.get_line_content(line_number)
                if (
                    len(line_content) == 0
                    or line_content == prev_content
                    or first_non_whitespace_index(line_content) != -1
                ):
                    # 如果行内容为空、与之前内容相同，或者行中存在非空白字符，跳过
                    continue
                trim_auto_whitespace_line_numbers.append(line_number)
        # self._trim_auto_whitespace_lines 是为了进行自动空格尾随的裁剪，但目前尚未启用
        self._trim_auto_whitespace_lines = trim_auto_whitespace_line_numbers
        return [TextEdit(range=reverse_op["range"], new_text=reverse_op["text"]) for reverse_op in reverse_operations]

    def _do_apply_edits(self, operations: list[OperationDict]) -> list[dict]:
        # 假设有一个静态方法来对操作进行降序排序
        operations.sort(key=lambda x: x["range"], reverse=True)

        content_changes = []
        for op in operations:
            start_line_number = op["range"].start_position.line
            start_column = op["range"].start_position.character
            end_line_number = op["range"].end_position.line
            end_column = op["range"].end_position.character

            if start_line_number == end_line_number and start_column == end_column and len(op["text"]) == 0:
                # no-op
                continue

            if op["text"]:
                # replacement
                self._delete(op["range"])
                self._insert(op["range"].start_position, op["text"])
            else:
                # deletion
                self._delete(op["range"])

            content_change_range = Range(
                start_position=Position(line=start_line_number, character=start_column),
                end_position=Position(line=end_line_number, character=end_column),
            )
            content_changes.append(
                {
                    "range": content_change_range,
                    "range_length": op["range_length"],
                    "text": op["text"],
                    "range_offset": op["range_offset"],
                    "force_move_markers": op["force_move_markers"],
                },
            )

        return content_changes

    def _delete(self, t_range: Range) -> None:
        """
        Delete the text in the given range.

        Args:
            t_range (Range): The range to delete.

        Returns:
            None
        """
        # 调整为0基索引
        start_line = t_range.start_position.line - 1
        end_line = t_range.end_position.line - 1
        start_column = t_range.start_position.character - 1
        end_column = t_range.end_position.character - 1

        if start_line == end_line:
            # 单行删除
            self._content[start_line] = (
                self._content[start_line][:start_column] + self._content[start_line][end_column:]
            )
        else:
            # 多行删除
            self._content[start_line] = self._content[start_line][:start_column]
            self._content[start_line] += self._content[end_line][end_column:]
            del self._content[start_line + 1 : end_line + 1]

    def _insert(self, position: Position, text: str) -> None:
        """
        Insert the text at the given position.

        Args:
            position (Position): The position to insert the text.
            text (str): The text to insert.

        Returns:
            None
        """
        # 变更为0基索引
        line_index: int = position.line - 1
        char_index: int = position.character - 1
        lines: list[str] = text.splitlines(False)  # 不保留换行符
        end_with_newline: bool = text.endswith("\n") or text.endswith("\r") or text.endswith("\r\n")
        if len(lines) == 1 and end_with_newline:
            # Python原生splitlines无法像split函数一样，在结尾遇到换行符的时候生成一个空行占位。所以在此需要判断
            lines.append("")

        if len(lines) == 1:
            # 单行插入
            self._content[line_index] = (
                self._content[line_index][:char_index] + lines[0] + self._content[line_index][char_index:]
            )
        else:
            # 多行插入
            start_line = self._content[line_index][:char_index] + lines[0]
            end_line = lines[-1] + self._content[line_index][char_index:]
            # 替换首尾行并插入新行
            self._content = (
                self._content[:line_index] + [start_line] + lines[1:-1] + [end_line] + self._content[line_index + 1 :]
            )

    @staticmethod
    def _get_inverse_edit_ranges(operations: list[OperationDict]) -> list[Range]:
        """
        Notes:
            此时需要operations按range倒序排列。因为是私有方法，所以默认调用者知道这个规矩，不再重复排序

        Args:
            operations: A list of operations representing the edits to be performed.

        Returns:
            A list of ranges representing the inverse edits.

        """
        result = []
        prev_op_end_line_number = 0
        prev_op_end_column = 0
        prev_op = None
        operations = copy.deepcopy(operations)

        operations.sort(key=lambda x: x["range"])

        for op in operations:
            if prev_op:
                if prev_op["range"].end_position.line == op["range"].start_position.line:
                    start_line_number = prev_op_end_line_number
                    start_column = prev_op_end_column + (
                        op["range"].start_position.character - prev_op["range"].end_position.character
                    )
                else:
                    start_line_number = prev_op_end_line_number + (
                        op["range"].start_position.line - prev_op["range"].end_position.line
                    )
                    start_column = op["range"].start_position.character
            else:
                start_line_number = op["range"].start_position.line
                start_column = op["range"].start_position.character

            if op["text"]:
                # the operation inserts something
                line_count = op["eol_count"] + 1
                if line_count == 1:
                    # single line insert
                    result_range = Range(
                        start_position=Position(line=start_line_number, character=start_column),
                        end_position=Position(
                            line=start_line_number,
                            character=start_column + op["first_line_length"],
                        ),
                    )
                else:
                    # multi line insert
                    result_range = Range(
                        start_position=Position(line=start_line_number, character=start_column),
                        end_position=Position(
                            line=start_line_number + line_count - 1,
                            character=op["last_line_length"] + 1,
                        ),
                    )
            else:
                # There is nothing to insert
                result_range = Range(
                    start_position=Position(line=start_line_number, character=start_column),
                    end_position=Position(line=start_line_number, character=start_column),
                )

            prev_op_end_line_number = result_range.end_position.line
            prev_op_end_column = result_range.end_position.character
            result.append(result_range)
            prev_op = op

        return result

    def insert_cursor(self, key: str, position: Position, strict: bool = False) -> str:
        """
        Insert a cursor at the given position.

        Args:
            key (str): The key of the cursor.
            position (Position): The position of the cursor.
            strict (bool): Whether to strictly insert the cursor. Defaults to False.

        Returns:
            str: The current view near the cursor.
        """
        self._assert_not_disposed()
        if self._is_valid_position(position):
            if any(c.key == key or c.position == position for c in self.cursors.values()) and strict:
                raise ValueError("Cursor already exists. Key and position must be unique.")
        else:
            if strict:
                raise ValueError("Invalid position. Please check the position's line and character.")
            else:
                position = self.validate_position(position)
        # 插入一个新的光标
        self.cursors[key] = Cursor(key=key, position=position)
        near_range = Range(
            start_position=Position(max(position.line - 5, 1), 1),
            end_position=Position(position.line + 5, 1),
        )
        return self.get_view(content_range=near_range)

    def delete_cursor(self, key: str) -> str:
        """
        Delete a cursor by the given key.

        Args:
            key (str): The key of the cursor.
        """
        self._assert_not_disposed()
        near_range: Range | None = None
        if key in self.cursors:
            near_range = Range(
                start_position=Position(max(self.cursors[key].position.line - 5, 1), 1),
                end_position=Position(self.cursors[key].position.line + 5, 1),
            )
            del self.cursors[key]
        return self.get_view(content_range=near_range)

    def clear_cursors(self, re_init: bool = False) -> str:
        """
        Clear all cursors.

        Args:
            re_init (bool): If re-initialize the cursors. Defaults to False.
        """
        self._assert_not_disposed()
        self.cursors.clear()
        if re_init:
            self.cursors["primary"] = Cursor(key="primary", position=Position(1, 1))
        return self.get_view()

    def get_simple_view(self, content_range: Range | None = None) -> str:
        """
        Get a simple view of the content.

        Args:
            content_range (Optional[Range]): The range of content to get. If None, the full content will be returned.

        Returns:
            str: The simple view of the content.
        """
        content_bak = copy.deepcopy(self._content)
        if not content_range:
            content_range = self.get_full_model_range()
        else:
            content_range = self.validate_range(content_range)
        try:
            contents = self.get_value_in_range(content_range)
            content_list = contents.splitlines(keepends=False)
            content_list = [
                f"{line_num: <5}|{line_content}"
                for line_num, line_content in enumerate(content_list, start=content_range.start_position.line)
            ]
            contents = self.get_eol().join(content_list)
            return self.simple_view_template.format(
                content_value=contents,
                uri=self.uri,
                language_id=self.language_id.value,
            )
        finally:
            self._content = content_bak

    def get_view(self, with_line_num: bool = True, content_range: Range | None = None) -> str:
        """
        获取当前视图，与get_value的区别在于，视图会带上光标信息与一些提示信息。方便大模型进行理解相对位置。

        Args:
            with_line_num (bool): 是否带有行号。默认为True。
            content_range (Optional[Range]): The range of content to get. If None, the full content will be returned.

        Returns:
            str: The current view.
        """
        content_bak = copy.deepcopy(self._content)
        if not content_range:
            content_range = self.get_full_model_range()
        else:
            content_range = self.validate_range(content_range)
        start_offset = self.get_offset_at(content_range.start_position)
        end_offset = self.get_offset_at(content_range.end_position)
        try:
            original_content = self.get_value_in_range(content_range)
            reversed_cursors = sorted(self.cursors.values(), key=lambda x: x.position, reverse=True)
            # 与编辑文档一样，需要倒序插入光标，因为插入光标的时候，如果有多个光标，后插入的光标会影响前面的光标位置
            for c in reversed_cursors:
                if c.position in content_range:
                    self._insert(c.position, repr(c))
                    end_offset += len(repr(c))
            start_pos = self.get_position_at(start_offset)
            end_pos = self.get_position_at(end_offset)
            contents = (
                self.add_line_num_to_contents(start_pos, end_pos)
                if with_line_num
                else self.get_value_in_range(Range(start_position=start_pos, end_position=end_pos))
            )
            language_id = self.language_id
            uri = self.uri
            return self.view_template.format(
                uri=uri,
                language_id=language_id.value,
                content_value=contents,
                original_content=original_content,
            )
        finally:
            self._content = content_bak

    def add_line_num_to_contents(self, from_pos: Position, to_pos: Position) -> str:
        """
        Add line numbers to the contents.

        Format is add line_num ahead of each line content. and insert 5 spaces width between line_num and line content. Note
        that line_num width included in 5 spaces.

        Args:
            from_pos: start position of the content.
            to_pos: end position of the content.

        Returns:
            str: The content with line numbers.
        """
        contents = self.get_value_in_range(Range(start_position=from_pos, end_position=to_pos))
        content_list = contents.splitlines(keepends=False)
        content_list = [
            f"{line_num: <5}|{line_content}" for line_num, line_content in enumerate(content_list, start=from_pos.line)
        ]
        return self.get_eol().join(content_list)

    def get_render(
        self,
        jinja: str,
        with_line_num: bool = True,
        content_range: Range | None = None,
        with_cursor: bool = True,
    ) -> str:
        """
        获取当前内容按Jinja模板渲染后的结果。

        Notes:
            Jinja模板中可以使用的变量有：
                - uri: 当前文档的URI
                - language_id: 当前文档的语言ID
                - content_value: 当前文档的内容
                - original_content: 当前文档的原始内容

        Args:
            jinja (str): Jinja模板字符串。
            with_line_num (bool): 是否带有行号。默认为True。
            content_range (Optional[Range]): The range of content to get. If None, the full content will be returned.
            with_cursor (bool): 是否带有光标信息。默认为True。

        Returns:
            str: The rendered content.
        """
        content_bak = copy.deepcopy(self._content)
        if not content_range:
            content_range = self.get_full_model_range()
        else:
            content_range = self.validate_range(content_range)
        start_offset = self.get_offset_at(content_range.start_position)
        end_offset = self.get_offset_at(content_range.end_position)
        try:
            original_content = self.get_value_in_range(content_range)
            if with_cursor:
                reversed_cursors = sorted(self.cursors.values(), key=lambda x: x.position, reverse=True)
                # 与编辑文档一样，需要倒序插入光标，因为插入光标的时候，如果有多个光标，后插入的光标会影响前面的光标位置
                for c in reversed_cursors:
                    if c.position in content_range:
                        self._insert(c.position, repr(c))
                        end_offset += len(repr(c))
            start_pos = self.get_position_at(start_offset)
            end_pos = self.get_position_at(end_offset)
            contents = (
                self.add_line_num_to_contents(start_pos, end_pos)
                if with_line_num
                else self.get_value_in_range(Range(start_position=start_pos, end_position=end_pos))
            )
            language_id = self.language_id
            uri = self.uri
            template = Template(jinja)
            return template.render(
                uri=uri,
                language_id=language_id.value,
                content_value=contents,
                original_content=original_content,
            )
        finally:
            self._content = content_bak

    def get_character_count_in_range(self, t_range: Range, eol: EndOfLinePreference | None) -> int:
        """
        TODO 此函数未经测试
        Get the character count in the given range.

        Args:
            t_range (Range): The range to get the character count.
            eol (Optional[EndOfLinePreference]): The end of line preference.

        Returns:
            int: The character count in the range.
        """
        self._assert_not_disposed()
        if not is_pure_basic_ascii(self._content):
            result = 0
            from_line_number = t_range.start_position.line
            to_line_number = t_range.end_position.line
            for line_number in range(from_line_number, to_line_number + 1):
                line_content = self.get_line_content(line_number)
                from_offset = t_range.start_position.character - 1 if line_number == from_line_number else 0
                to_offset = t_range.end_position.character - 1 if line_number == to_line_number else len(line_content)
                offset = from_offset
                while offset < to_offset:
                    if self._is_high_surrogate(line_content[offset]):
                        result += 1
                        offset += 2
                    else:
                        result += 1
                        offset += 1
            result += len(self.get_eol_by_preference(eol)) * (to_line_number - from_line_number)
            return result
        return self.get_value_length_in_range(t_range, eol)

    @staticmethod
    def _is_high_surrogate(char: str) -> bool:
        """Check if the character is a high surrogate in the UTF-16 encoding."""
        return 0xD800 <= ord(char) <= 0xDBFF

    def is_attached_to_editor(self) -> bool:
        raise NotImplementedError

    def on_did_change_content(self, listener: Callable[[ModelContentChangedEvent], None]) -> DisposableProtocol:
        raise NotImplementedError
