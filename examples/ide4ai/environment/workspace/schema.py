# filename: schema.py
# @Time    : 2024/4/18 11:40
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import dataclasses
import struct
from collections.abc import Callable, Iterator
from enum import Enum
from functools import total_ordering
from typing import (
    Annotated,
    Any,
    Literal,
    NamedTuple,
    NoReturn,
    Optional,
    Protocol,
    Self,
    TypeAlias,
    runtime_checkable,
)

from annotated_types import Gt
from pydantic import AnyUrl, BaseModel, Field, GetJsonSchemaHandler, PrivateAttr, model_validator
from pydantic.json_schema import JsonSchemaValue
from pydantic_core import core_schema as cs

from ide4ai.dtos.text_documents import (
    LSPPosition,
    LSPRange,
    LSPTextEdit,
)
from ide4ai.environment.workspace.common.dispose import (
    DisposableProtocol,
)
from ide4ai.schema import LanguageId


@total_ordering
class Position(NamedTuple):
    line: Annotated[int, Gt(0)]
    character: Annotated[int, Gt(0)]

    def __lt__(self, other: Any) -> bool:
        if not isinstance(other, Position):
            return NotImplemented  # 让Python处理不兼容的类型比较 | Let Python handle incompatible type comparisons
        if self.line == other.line:
            return self.character < other.character
        return self.line < other.line

    def __gt__(self, other: Any) -> bool:
        if not isinstance(other, Position):
            return NotImplemented  # 让Python处理不兼容的类型比较 | Let Python handle incompatible type comparisons
        if self.line == other.line:
            return self.character > other.character
        return self.line > other.line

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, Position):
            return False  # pragma: no cover
        return self.line == other.line and self.character == other.character

    @classmethod
    def from_lsp_position(cls, lsp_position: LSPPosition) -> "Position":
        return cls(line=lsp_position.line + 1, character=lsp_position.character + 1)  # pragma: no cover

    def to_lsp_position(self) -> LSPPosition:
        return LSPPosition(line=self.line - 1, character=self.character - 1)

    @classmethod
    def __get_pydantic_json_schema__(cls, core_schema: cs.CoreSchema, handler: GetJsonSchemaHandler) -> JsonSchemaValue:
        """
        增强Pydantic 获取 JsonSchema的方法，将Position的line和character字段的类型改为integer

        如果不带上item的类型说明，GPT API会报错

        Args:
            core_schema:
            handler:

        Returns:

        """
        sch = handler(core_schema)
        json_schema = handler.resolve_ref_schema(sch)
        json_schema["items"] = {"type": "integer"}
        return json_schema


class Range(BaseModel, validate_assignment=True):
    """
    注意Range是从(1,1)开始的，区别于数组从(0,0)开始
    """

    start_position: Position = Field(
        title="开始位置",
        description="开始位置坐标，格式为 (line, character)，注意不可以使用'-1'这种负数表示倒数形式",
    )
    end_position: Position = Field(
        title="结束位置",
        description="结束位置坐标，格式为 (line, character)，注意不可以使用'-1'这种负数表示倒数形式",
    )

    @model_validator(mode="after")
    def validate_position(self) -> Self:
        """
        Validate the position of the range. start_position must be less than or equal to end_position.

        Returns:
            Self: The validated range.
        """
        assert self.start_position <= self.end_position, "开始位置必须小于等于结束位置"
        return self

    @classmethod
    def from_lsp_range(cls, lsp_range: LSPRange) -> "Range":
        return cls(  # pragma: no cover
            start_position=Position.from_lsp_position(lsp_range.start),
            end_position=Position.from_lsp_position(lsp_range.end),
        )

    def to_lsp_range(self) -> LSPRange:
        return LSPRange(
            start=self.start_position.to_lsp_position(),
            end=self.end_position.to_lsp_position(),
        )

    def is_empty(self) -> bool:
        """
        Test if this range is empty.

        Returns:
            bool: True if the range is empty, False otherwise.
        """
        return self.start_position == self.end_position

    def contains_position(self, position: Position) -> bool:
        """
        Test if this range contains the given position.

        Args:
            position (Position): The position to test.

        Returns:
            bool: True if the range contains the position, False otherwise.
        """
        return self.start_position <= position <= self.end_position

    def contains_range(self, t_range: "Range") -> bool:
        """
        Test if this range contains the given range.

        Args:
            t_range (Range): The range to test.

        Returns:
            bool: True if the range contains the given range, False otherwise.
        """
        return (self & t_range) == t_range

    def strict_contains_range(self, t_range: "Range") -> bool:
        """
        Test if this range contains the given range.

        Args:
            t_range (Range): The range to test.

        Returns:
            bool: True if the range contains the given range, False otherwise.
        """
        return self.start_position < t_range.start_position and self.end_position > t_range.end_position

    def plus_range(self, t_range: "Range") -> "Range":
        """
        Get the range of the union of this range and the given range.

        Args:
            t_range (Range): The range to union with.

        Returns:
            Range: The union of the two ranges.
        """
        return self | t_range

    def intersect_ranges(self, t_range: "Range") -> Optional["Range"]:
        """
        Get the range of the intersection of this range and the given range.

        Args:
            t_range (Range): The range to intersect with.

        Returns:
            Optional[Range]: The intersection of the two ranges, or None if there is no intersection.
        """
        return self & t_range

    def collapse_to_start(self) -> "Range":
        """
        Collapse this range to its start position.

        Returns:
            Range: The range collapsed to its start position.
        """
        return Range(start_position=self.start_position, end_position=self.start_position)

    def collapse_to_end(self) -> "Range":
        """
        Collapse this range to its end position.

        Returns:
            Range: The range collapsed to its end position.
        """
        return Range(start_position=self.end_position, end_position=self.end_position)

    def spans_multiple_lines(self) -> bool:
        """
        Test if this range spans multiple lines.

        Returns:
            bool: True if the range spans multiple lines, False otherwise.
        """
        return self.start_position.line != self.end_position.line

    def __lt__(self, other: "Range") -> bool:
        return self.start_position < other.start_position

    def __gt__(self, other: "Range") -> bool:
        return self.start_position > other.start_position

    def __eq__(self, other: object) -> bool:
        return (
            isinstance(other, Range)
            and self.start_position == other.start_position
            and self.end_position == other.end_position
        )

    def __repr__(self) -> str:
        return f"Range({self.start_position}, {self.end_position})"

    def __or__(self, other: "Range") -> "Range":
        new_start = min(self.start_position, other.start_position)
        new_end = max(self.end_position, other.end_position)
        return Range(start_position=new_start, end_position=new_end)

    def __ior__(self, other: "Range") -> Self:
        union = self | other
        self.start_position, self.end_position = (
            union.start_position,
            union.end_position,
        )
        return self

    def __and__(self, other: "Range") -> Optional["Range"]:
        new_start = max(self.start_position, other.start_position)
        new_end = min(self.end_position, other.end_position)
        if new_start > new_end:
            return None
        return Range(start_position=new_start, end_position=new_end)

    def __iand__(self, other: "Range") -> Self:
        intersection = self & other
        if intersection is None:
            raise ValueError("交集为空，无法更新")
        self.start_position, self.end_position = (
            intersection.start_position,
            intersection.end_position,
        )
        return self

    def __sub__(self, other: Any) -> list["Range"]:
        intersection = self & other
        if intersection is None:
            return [self]
        result = []
        if self.start_position < intersection.start_position:
            result.append(
                Range(
                    start_position=self.start_position,
                    end_position=Position(
                        intersection.start_position.line,
                        intersection.start_position.character,
                    ),
                ),
            )
        if self.end_position > intersection.end_position:
            result.append(
                Range(
                    start_position=Position(
                        intersection.end_position.line,
                        intersection.end_position.character,
                    ),
                    end_position=self.end_position,
                ),
            )
        return result

    def __isub__(self, other: Any) -> Self:
        difference = self - other
        if not difference:
            raise ValueError("结果为空，无法更新")
        elif len(difference) == 1:
            self.start_position, self.end_position = (
                difference[0].start_position,
                difference[0].end_position,
            )
        else:
            raise ValueError("差集操作结果为多个区间，无法更新")
        return self

    def __add__(self, other: Any) -> NoReturn:
        raise NotImplementedError("加法不适用于区间，请使用并集（| 或 |=）或交集（& 或 &=）")

    def __iadd__(self, other: Any) -> NoReturn:
        raise NotImplementedError("加法不适用于区间，请使用并集（| 或 |=）或交集（& 或 &=）")

    def __contains__(self, item: Any) -> bool:
        """
        判断某个Range或者Position是否在Range内部
        Args:
            item (Any): 合法的Position或者Range

        Returns:
            bool: 是否在Range内部
        """
        if isinstance(item, Range):
            return self.start_position <= item.start_position and self.end_position >= item.end_position
        elif isinstance(item, Position):
            return self.start_position <= item <= self.end_position
        else:
            raise ValueError("不支持的类型")


class SearchResult(BaseModel):
    """
    SearchResult contains the result of a search operation.
    """

    # The range of the search result.
    range: Range
    # The matching text if capture_matches is True.
    match: str | None = None

    def __repr__(self) -> str:
        return f"SearchResult: {self.match}.\n\tRange Scope: {repr(self.range)}"


class SingleEditOperation(BaseModel):
    """
    A single edit operation, that acts as a simple replace. i.e. Replace text at range with text in model.
    """

    range: Range = Field(
        title="编辑范围",
        description="The range to replace. This can be empty to emulate a simple insert.",
    )
    text: str | None = Field(
        title="替换文本",
        description="The text to replace with. This can be empty to emulate a simple delete.",
    )
    # This indicates that this operation has "insert" semantics. i.e. forceMoveMarkers = true => if range is collapsed,
    # all markers at the position will be moved. | 这个操作具有"插入"的语义。也就是说，当 forceMoveMarkers = true 时，如果范围
    # （range）已经折叠（即开始和结束位置相同），那么所有位于该位置的标记都将被移动。
    force_move_markers: bool | None = Field(
        default=False,
        title="强制移动标记",
        description="Whether to force moving the cursor markers, even if the edit is a no-op.",
    )


class IdentifiedSingleEditOperation(SingleEditOperation):
    """
    A single edit operation with an identifier. i.e. Replace text at range with text in model.
    """

    # TODO: 如果注释掉text与force_move_markers字段，会引起mypy的递归判断，这个问题比较奇怪，可能是mypy的问题。需要进一步确认
    text: str | None = Field(
        title="替换文本",
        description="The text to replace with. This can be empty to emulate a simple delete.",
    )
    # This indicates that this operation has "insert" semantics. i.e. forceMoveMarkers = true => if range is collapsed,
    # all markers at the position will be moved. | 这个操作具有"插入"的语义。也就是说，当 forceMoveMarkers = true 时，如果范围
    # （range）已经折叠（即开始和结束位置相同），那么所有位于该位置的标记都将被移动。
    force_move_markers: bool | None = Field(
        default=False,
        title="强制移动标记",
        description="Whether to force moving the cursor markers, even if the edit is a no-op.",
    )

    identifier: str = Field(
        title="操作标识符",
        description="An identifier for the edit operation. This identifier is used to track the edit operation.",
    )
    is_auto_whitespace_edit: bool = Field(default=False, title="是否自动空格编辑")
    is_tracked: bool = Field(default=False, title="是否跟踪")


class FindMatch(BaseModel):
    range: Range  # Specify the correct type based on what 'range' is expected to be in your application
    matches: list[Any]  # Specify the correct type based on what 'matches' typically contains
    _findMatchBrand: Any | None = PrivateAttr(
        None,
    )  # Use 'Optional' to indicate this can be undefined, similar to TypeScript's 'undefined'


class Cursor(BaseModel):
    key: str = Field(
        title="光标名称Key",
        description="光标在某个Model中的只一标识符",
    )
    position: Position = Field(
        default_factory=lambda: Position(0, 0),
        title="光标位置",
        description="光标的位置坐标，格式为 (line, character)，注意不可以使用'-1'这种负数表示倒数形式",
    )

    def __repr__(self) -> str:
        """
        使用单冒号对光标进行格式化输出

        Returns:
            str: 格式化后的光标输出
        """
        return f">{self.key}|{self.position.line}:{self.position.character}<"


class TextEdit(BaseModel):
    # The range of the text document to be manipulated. To insert text into a document create a range where
    # start === end.
    range: Range = Field(
        title="编辑范围",
        description="The range of the text to edit.",
    )
    # The string to be inserted. For delete operations use an empty string.
    new_text: str = Field(
        title="替换文本",
        description="The text to replace with.",
    )

    def to_lsp_text_edit(self) -> "LSPTextEdit":
        return LSPTextEdit(
            range=self.range.to_lsp_range(),
            new_text=self.new_text,
        )

    def to_single_edit_operation(self) -> SingleEditOperation:
        return SingleEditOperation(
            range=self.range,
            text=self.new_text,
        )


def escape_new_line(s: str) -> str:
    """Escapes new line characters in a string for display."""
    return s.replace("\n", "\\n").replace("\r", "\\r")


class TextChange(BaseModel):
    old_position_offset: int
    old_text: str
    new_position_offset: int
    new_text: str

    @property
    def old_length(self) -> int:
        """Returns the length of the old text."""
        return len(self.old_text)

    @property
    def old_end(self) -> int:
        """Calculates the end position of the old text."""
        return self.old_position_offset + len(self.old_text)

    @property
    def new_length(self) -> int:
        """Returns the length of the new text."""
        return len(self.new_text)

    @property
    def new_end(self) -> int:
        """Calculates the end position of the new text."""
        return self.new_position_offset + len(self.new_text)

    def __str__(self) -> str:
        """Returns a string representation of the text change."""
        if self.old_text == "":
            return f'(insert@{self.old_position_offset} "{escape_new_line(self.new_text)}")'
        if self.new_text == "":
            return f'(delete@{self.old_position_offset} "{escape_new_line(self.old_text)}")'
        return f'(replace@{self.old_position_offset} "{escape_new_line(self.old_text)}" with "{escape_new_line(self.new_text)}")'

    def write(self, buffer: bytearray, offset: int) -> int:
        """Writes the TextChange to a bytearray starting from the given offset."""
        old_text_encoded = self.old_text.encode("utf-16le")
        new_text_encoded = self.new_text.encode("utf-16le")

        # Calculate the required size
        required_size = offset + 4 + 4 + (4 + len(old_text_encoded)) + (4 + len(new_text_encoded))

        # Extend the buffer if necessary
        if len(buffer) < required_size:
            buffer.extend(bytearray(required_size - len(buffer)))

        # Write old_position and new_position
        struct.pack_into(">I", buffer, offset, self.old_position_offset)
        offset += 4
        struct.pack_into(">I", buffer, offset, self.new_position_offset)
        offset += 4

        # Write old_text
        offset = self._write_string(buffer, old_text_encoded, offset)

        # Write new_text
        offset = self._write_string(buffer, new_text_encoded, offset)

        return offset

    @staticmethod
    def _write_string(buffer: bytearray, data: bytes, offset: int) -> int:
        """Writes a string's length and data to the buffer."""
        struct.pack_into(">I", buffer, offset, len(data))
        offset += 4
        buffer[offset : offset + len(data)] = data
        offset += len(data)
        return offset

    @staticmethod
    def read(buffer: bytes, offset: int) -> tuple["TextChange", int]:
        """Reads a TextChange from a bytes buffer starting from the given offset."""
        (old_position,) = struct.unpack_from(">I", buffer, offset)
        offset += 4
        (new_position,) = struct.unpack_from(">I", buffer, offset)
        offset += 4
        old_text, offset = TextChange._read_string(buffer, offset)
        new_text, offset = TextChange._read_string(buffer, offset)

        return (
            TextChange(
                old_position_offset=old_position,
                old_text=old_text,
                new_position_offset=new_position,
                new_text=new_text,
            ),
            offset,
        )

    @staticmethod
    def _read_string(buffer: bytes, offset: int) -> tuple[str, int]:
        """Reads a string from the buffer."""
        (length,) = struct.unpack_from(">I", buffer, offset)
        offset += 4
        text = buffer[offset : offset + length].decode("utf-16le")
        offset += length
        return text, offset


@dataclasses.dataclass(frozen=True)
class ModelContentChange:
    # The range that got replaced.
    range: Range
    # The offset of the range that got replaced.
    range_offset: int
    # The length of the range that got replaced.
    range_length: int
    # The new text for the range.
    text: str


@dataclasses.dataclass(frozen=True)
class ModelContentChangedEvent:
    # The changes to the model content.
    changes: list[ModelContentChange]
    # The (new) end-of-line character.
    eol: str
    # The new version id the model has transitioned to.
    version_id: int
    # Flag that indicates that this event was generated while undoing.
    is_undoing: bool
    # Flag that indicates that this event was generated while redoing.
    is_redoing: bool
    # Flag that indicates that all decorations were lost with this edit.
    is_flush: bool
    # Flag that indicates that this event was generated while undoing the last edit.
    is_eol_change: bool


CursorStateComputer: TypeAlias = Callable[[list[Range] | None], list[Range] | None]


class DefaultEndOfLine(Enum):
    # Use line feed (\n) as the end of line character.
    LF = 1
    # Use carriage return and line feed (\r\n) as the end of line character.
    CRLF = 2


class EndOfLinePreference(Enum):
    # Use the end of line character identified in the text buffer.
    TEXT_DEFINED = 0
    # Use line feed (\n) as the end of line character.
    LF = 1
    # Use carriage return and line feed (\r\n) as the end of line character.
    CRLF = 2


class EndOfLineSequence(Enum):
    # Use line feed (\n) as the end of line character.
    LF = 0
    # Use carriage return and line feed (\r\n) as the end of line character.
    CRLF = 1


@dataclasses.dataclass(frozen=True)
class TextModelResolvedOptions:
    tab_size: int
    indent_size: int
    insert_spaces: bool
    default_eol: DefaultEndOfLine
    trim_auto_whitespace: bool
    _text_model_resolved_options_brand: Callable | None = None

    @property
    def original_indent_size(self) -> int | Literal["tabSize"]:
        raise NotImplementedError


@dataclasses.dataclass(frozen=True)
class WordAtPosition:
    word: str
    start_column: int
    end_column: int


@runtime_checkable
class ModelProtocol(Protocol):
    """
    ModelProtocol is a protocol for model.
    """

    @property
    def uri(self) -> AnyUrl:
        """
        The associated URI for this document. Most documents have the file-scheme, indicating that they represent files
        on disk. However, some documents may have other schemes indicating that they are not available on disk.

        Returns:

        """
        raise NotImplementedError

    @property
    def m_id(self) -> str:
        """
        The ID of the model.

        Returns:
            str: The ID of the model.
        """
        raise NotImplementedError

    def get_options(self) -> TextModelResolvedOptions:
        """
        Get the resolved options for this model.

        Returns:
            TextModelResolvedOptions: The resolved options for this model.
        """
        ...

    def get_version_id(self) -> int:
        """
        Get the current version id of the model. Anytime a change happens to the model (even undo/redo), the version id
        is incremented.

        Returns:
            int: The version id of the model.
        """
        ...

    def get_alternative_version_id(self) -> int:
        """
        Get the alternative version id of the model. This alternative version id is not always incremented, it will
        return the same values in the case of undo-redo.

        Returns:
            int: The alternative version id of the model.
        """
        ...

    def set_value(self, new_value: str | Iterator[str]) -> None:
        """
        Replace the entire text buffer value contained in this model.

        Args:
            new_value (str | Iterator[str]): The new value to set.
        """
        ...

    def get_value(
        self,
        eol: EndOfLinePreference | None = None,
        preserve_bom: bool | None = None,
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
        ...

    def create_snapshot(self, preserve_bom: bool | None = None) -> Iterator[str]:
        """
        Get the text stored in this model.

        Args:
            preserve_bom (bool): Preserve a BOM character if it was detected when the model was constructed.

        Returns:
            Iterator[str]: The text stored in this model.
        """
        ...

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
        ...

    def get_value_in_range(self, t_range: Range, eol: EndOfLinePreference | None) -> str:
        """
        Get the text stored in this model by the given range.

        Args:
            t_range (Range): The range of the text.
            eol (EndOfLinePreference): The end of line character preference. This will only be used for multiline
                ranges. Defaults to EndOfLinePreference.TextDefined.

        Returns:
            str: The text stored in this model.
        """
        ...

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
        ...

    def get_character_count_in_range(self, t_range: Range, eol: EndOfLinePreference | None) -> int:
        """
        Get the character count in the given range.

        Args:
            t_range (Range): The range to count characters.
            eol (EndOfLinePreference): The end of line character preference. This will only be used for multiline
                ranges. Defaults to EndOfLinePreference.TextDefined.

        Returns:
            int: The character count in the given range.
        """
        ...

    def get_line_count(self) -> int:
        """
        Get the number of lines in the model.

        Returns:
            int: The number of lines in the model.
        """
        ...

    def get_line_content(self, line_number: int) -> str:
        """
        Get the content of a line in the model.

        Args:
            line_number (int): The line number.

        Returns:
            str: The content of the line.
        """
        ...

    def get_line_length(self, line_number: int) -> int:
        """
        Get the length of a line in the model.

        Args:
            line_number (int): The line number.

        Returns:
            int: The length of the line.
        """
        ...

    def get_lines_content(self) -> list[str]:
        """
        Get the content of all lines in the model.

        Returns:
            list[str]: The content of all lines in the model.
        """
        ...

    def get_eol(self) -> str:
        """
        Get the end of line character in the model.

        Returns:
            str: The end of line character in the model.
        """
        ...

    def get_end_of_line_sequence(self) -> EndOfLineSequence:
        """
        Get the end of line sequence in the model.

        Returns:
            EndOfLineSequence: The end of line sequence in the model.
        """
        ...

    def get_line_min_column(self, line_num: int) -> int:
        """
        Get the minimum column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The minimum column of the line.
        """
        ...

    def get_line_max_column(self, line_num: int) -> int:
        """
        Get the maximum column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The maximum column of the line.
        """
        ...

    def get_line_first_non_whitespace_column(self, line_num: int) -> int:
        """
        Get the first non-whitespace column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The first non-whitespace column of the line.
        """
        ...

    def get_line_last_non_whitespace_column(self, line_num: int) -> int:
        """
        Get the last non-whitespace column of a line.

        Args:
            line_num (int): The line number.

        Returns:
            int: The last non-whitespace column of the line.
        """
        ...

    def validate_position(self, position: Position) -> Position:
        """
        Validate the given position.

        Args:
            position (Position): The position to validate.

        Returns:
            Position: The validated position.
        """
        ...

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
        ...

    def validate_range(self, t_range: Range) -> Range:
        """
        Validate the given range.

        Args:
            t_range (Range): The range to validate.

        Returns:
            Range: The validated range.
        """
        ...

    def get_offset_at(self, position: Position) -> int:
        """
        Get the offset at the given position.

        Args:
            position (Position): The position to get the offset.

        Returns:
            int: The offset at the given position.
        """
        ...

    def get_position_at(self, offset: int) -> Position:
        """
        Get the position at the given offset.

        Args:
            offset (int): The offset to get the position.

        Returns:
            Position: The position at the given offset.
        """
        ...

    def get_full_model_range(self) -> Range:
        """
        Get a range covering the entire model.

        Returns:
            Range: The full model range.
        """
        ...

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
        Find all matches of the search string in the model.

        Args:
            search_string (str): The search string.
            search_scope (Range | list[Range] | None): The search mode.
            is_regex (bool): Whether the search string is a regex.
            match_case (bool): Whether to match case.
            word_separator (Optional[str]): The word separator.
            capture_matches (bool): Whether to capture matches.
            limit_result_count (Optional[int]): The limit of the result count.

        Returns:
            list[Range]: The list of ranges of all matches.
        """
        ...

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
        """
        Find the next match of the search string in the model.

        Args:
            search_string (str): The search string.
            search_start (Position): The search start position.
            search_mode (bool | Range | list[Range]): The search mode.
            is_regex (bool): Whether the search string is a regex.
            match_case (bool): Whether to match case.
            word_separator (Optional[str]): The word separator.
            capture_matches (bool): Whether to capture matches.

        Returns:
            Optional[Range]: The range of the next match, or None if no match is found.
        """
        ...

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
        """
        Find the previous match of the search string in the model.

        Args:
            search_string (str): The search string.
            search_start (Position): The search start position.
            search_mode (bool | Range | list[Range]): The search mode.
            is_regex (bool): Whether the search string is a regex.
            match_case (bool): Whether to match case.
            word_separator (Optional[str]): The word separator.
            capture_matches (bool): Whether to capture matches.

        Returns:
            Optional[Range]: The range of the previous match, or None if no match is found.
        """
        ...

    def get_language_id(self) -> str:
        """
        Get the language id of the model.

        Returns:
            str: The language id of the model.
        """
        ...

    def get_word_at_position(self, position: Position) -> WordAtPosition | None:
        """
        Get the word under or besides position.

        Args:
            position (Position): The position to get the word.

        Returns:
            Optional[WordAtPosition]: The word at the given position, or None if no word is found.
        """
        ...

    def get_word_until_position(self, position: Position) -> WordAtPosition:
        """
        Get the word until position.

        Args:
            position (Position): The position to get the word.

        Returns:
            WordAtPosition: The word until the given position.
        """
        ...

    def update_options(self, options: TextModelResolvedOptions) -> None:
        """
        Update the options of the model.

        Args:
            options (TextModelResolvedOptions): The options to update.
        """
        ...

    def normalize_indentation(self, text: str) -> str:
        """
        Normalize a string containing whitespace according to indentation rules (converts to spaces or to tabs).

        Args:
            text (str): The text to normalize.

        Returns:
            str: The normalized text.
        """
        ...

    def detect_indentation(self, default_insert_spaces: bool, default_tab_size: int) -> None:
        """
        Detect the indentation options for this model from its content.

        Args:
            default_insert_spaces (bool): The default insert spaces option.
            default_tab_size (int): The default tab size.
        """
        ...

    def push_stack_element(self) -> None:
        """
        Close the current undo-redo element. This offers a way to create an undo/redo stop point.
        """
        ...

    def pop_stack_element(self) -> None:
        """
        Open the current undo-redo element. This offers a way to remove the current undo/redo stop point.
        """
        ...

    def push_edit_operations(
        self,
        before_cursor_state: list[Range] | None,
        edit_operations: list[SingleEditOperation],
        cursorStateComputer: CursorStateComputer,
    ) -> list[Range] | None:
        """
        Push edit operations, basically editing the model. This is the preferred way of editing the model. The edit
        operations will land on the undo stack.

        Args:
            before_cursor_state (Optional[list[Range]]): The cursor state before the edit operations. This cursor state
                will be returned when undo or redo are invoked.
            edit_operations (list[IdentifiedSingleEditOperation]): The edit operations to push.
            cursorStateComputer (CursorStateComputer): A callback that can compute the resulting cursors state after the
                edit operations have been executed.

        Returns:
            Optional[list[Range]]: The state of the cursor after the edit.
        """
        ...

    def push_eol(self, eol: EndOfLineSequence) -> None:
        """
        Change the end of line sequence. This is the preferred way of changing the eol sequence. This will land on the
        undo stack.

        Args:
            eol (EndOfLineSequence): The end of line sequence to push.
        """
        ...

    def apply_edits(
        self,
        operations: list[SingleEditOperation],
        computeUndoEdits: bool | None = None,
    ) -> list[TextEdit] | None:
        """
        Edit the model without adding the edits to the undo stack. This can have dire consequences on the undo stack!
        See @push_edit_operations for the preferred way.

        Args:
            operations (list[IdentifiedSingleEditOperation]): The operations to apply.
            computeUndoEdits (Optional[bool]): Whether to compute undo edits.

        Returns:
            Optional[list[TextEdit]]: The text edits. If desired, the inverse edit operations, that, when applied,
                will bring the model back to the previous state.
        """
        ...

    def set_eol(self, eol: EndOfLineSequence) -> None:
        """
        Change the end of line sequence without recording in the undo stack. This can have dire consequences on the
        undo stack! See @pushEOL for the preferred way.

        Args:
            eol:

        Returns:
            None
        """
        ...

    def on_did_change_content(self, listener: Callable[[ModelContentChangedEvent], None]) -> DisposableProtocol:
        """
        Register a listener for content changes.

        Args:
            listener (Callable[[ModelContentChangedEvent], None]): The listener to register.

        Returns:
            DisposableProtocol: A disposable to unregister the listener.
        """
        ...

    def is_attached_to_editor(self) -> bool:
        """
        Returns if this model is attached to an editor or not.

        Returns:
            bool: True if the model is attached to an editor, False otherwise.
        """
        ...

    def dispose(self) -> None:
        """
        Destroy the model.

        Returns:
            None
        """
        ...

    def insert_cursor(self, key: str, position: Position) -> str:
        """
        Insert a cursor at the given position.

        Args:
            key (str): The key of the cursor.
            position (Position): The position of the cursor.

        Returns:
            str: 修改后的视图内容
        """
        ...

    def delete_cursor(self, key: str) -> str:
        """
        Delete a cursor by the given key.

        Args:
            key (str): The key of the cursor.

        Returns:
            str: 修改后的视图内容
        """
        ...

    def clear_cursors(self) -> str:
        """
        Clear all cursors.

        Returns:
            str: 修改后的视图内容
        """
        ...

    def get_view(self) -> str:
        """
        Get the view of the model.

        获取当前视图，与get_value的区别在于，视图会带上光标信息与一些提示信息。方便大模型进行理解相对位置。

        Returns:
            str: The view of the model.
        """
        ...


class TextDocumentProtocol(Protocol):
    """
    A simple text document. The document keeps the content as string.
    """

    @property
    def uri(self) -> AnyUrl:
        """
        The associated URI for this document. Most documents have the file-scheme, indicating that they represent files
        on disk. However, some documents may have other schemes indicating that they are not available on disk.

        Returns:

        """
        raise NotImplementedError

    @property
    def language_id(self) -> LanguageId:
        """
        The identifier of the language associated with this document.

        Returns:
            str: The language identifier.
        """
        raise NotImplementedError

    @property
    def version(self) -> int:
        """
        The version number of this document (it will strictly increase after each change, including undo/redo).

        Returns:
            int: The version number of this document.
        """
        raise NotImplementedError

    @property
    def line_count(self) -> int:
        """
        The number of lines in this document.

        Returns:
            int: The number of lines in this document.
        """
        raise NotImplementedError

    def get_text(self, t_range: Range | None) -> str:
        """
        Get the text of this document. A substring can be retrieved by providing a range.

        Args:
            t_range (Optional[Range]): (optional) An range within the document to return. If no range is passed, the
                full content is returned. Invalid range positions are adjusted as described in Position.line and
                Position.character.
                If the start range position is greater than the end range position, then the effect of getText is as
                if the two positions were swapped.

        Returns:
            str: The text of this document or a substring of the text if a range is provided.
        """
        ...

    def position_at(self, offset: int) -> Position:
        """
        Get the position at the given offset.

        Args:
            offset (int): The offset to get the position.

        Returns:
            Position: The position at the given offset.
        """
        ...

    def offset_at(self, position: Position) -> int:
        """
        Get the offset at the given position.

        Args:
            position (Position): The position to get the offset.

        Returns:
            int: The offset at the given position.
        """
        ...


class TextDocumentItemProtocol(Protocol):
    """
    An item to transfer a text document from the client to the server.
    """

    uri: AnyUrl  # The text document's URI.
    language_id: str  # The text document's language identifier.
    version: int  # The version number of this document.
    text: str  # The content of the opened text document.
