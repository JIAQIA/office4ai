# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.

from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel, Field, PrivateAttr

from ide4ai.environment.workspace.schema import Range


# Vertical Lane in the overview ruler of the editor
class OverviewRulerLane(Enum):
    Left = 1
    Center = 2
    Right = 4
    Full = 7


# Vertical Lane in the glyph margin of the editor
class GlyphMarginLane(Enum):
    Left = 1
    Center = 2
    Right = 3


class InjectedTextCursorStops(Enum):
    Both = 0
    Right = 1
    Left = 2
    None_ = 3


class TextModelResolvedOptions(BaseModel):
    _text_model_resolved_options_brand: Any = PrivateAttr(None)
    tab_size: int
    indent_size: int
    _indent_size_is_tab_size: bool = PrivateAttr(False)
    insert_spaces: bool
    default_eol: str
    trim_auto_whitespace: bool
    bracket_pair_colorization_options: dict = Field(default_factory=dict)

    @property
    def original_indent_size(self) -> Literal["tab_size"] | int:
        return "tab_size" if self._indent_size_is_tab_size else self.indent_size

    def equals(self, other: Any) -> bool:
        return (
            isinstance(other, TextModelResolvedOptions)
            and self.tab_size == other.tab_size
            and self.indent_size == other.indent_size
            and self.insert_spaces == other.insert_spaces
            and self.default_eol == other.default_eol
            and self.trim_auto_whitespace == other.trim_auto_whitespace
            and self.bracket_pair_colorization_options == other.bracket_pair_colorization_options
        )

    def create_change_event(self, new_opts: Any) -> dict:
        return {
            "tab_size": self.tab_size != new_opts.tab_size,
            "indent_size": self.indent_size != new_opts.indent_size,
            "insert_spaces": self.insert_spaces != new_opts.insert_spaces,
            "trim_auto_whitespace": self.trim_auto_whitespace != new_opts.trim_auto_whitespace,
        }


class FindMatch(BaseModel):
    _find_match_brand: Any = PrivateAttr(None)
    range: Range
    matches: list


def is_text_snapshot(obj: Any) -> bool:
    return hasattr(obj, "read")


class ValidAnnotatedEditOperation(BaseModel):
    identifier: dict
    range: tuple
    text: str
    force_move_markers: bool
    is_auto_whitespace_edit: bool
    is_tracked: bool


class SearchData(BaseModel):
    regex: str
    word_separators: str
    simple_search: bool


class ApplyEditsResult(BaseModel):
    reverse_edits: list
    changes: list
    trim_auto_whitespace_line_numbers: list
