# filename: model_search.py
# @Time    : 2024/5/6 12:27
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import re
from typing import Any, ClassVar

from pydantic import BaseModel, ConfigDict

from ide4ai.environment.workspace.common.core.word_character_classifier import (
    WordCharacterClassifier,
    get_map_for_word_separators,
)
from ide4ai.environment.workspace.model import TextModel
from ide4ai.environment.workspace.schema import (
    EndOfLinePreference,
    Position,
    Range,
    SearchResult,
)


def escape_regexp_characters(input_string: str) -> str:
    """
    Escapes all regex characters in the string to treat them as literal characters.
    """
    return re.escape(input_string)


class SearchData(BaseModel):
    """
    SearchData contains the regex and other data needed to perform a search.
    """

    # The regex pattern to search for.
    regex: re.Pattern
    # The word separators to use when searching.
    word_separators: WordCharacterClassifier | None = None
    # The original search string if simple search is used.
    simple_search: str | None = None

    model_config: ClassVar[ConfigDict] = ConfigDict(arbitrary_types_allowed=True)


class SearchParams(BaseModel):
    """
    SearchParams manage the parameters and processing of a search query.
    """

    # The search query as a string.
    search_string: str
    # Whether the search query is a regex.
    is_regex: bool = False
    # Whether the search should be case-sensitive.
    match_case: bool = False
    # Optional string of word separators.
    word_separators: str | None = None

    def parse_search_request(self) -> SearchData | None:
        if not self.search_string:
            return None  # pragma: no cover

        # Try to create a regex out of the params
        multiline: bool
        if self.is_regex:
            multiline = self.is_multiline_regex_source()
        else:
            multiline = "\n" in self.search_string

        try:
            regex = self.create_regexp()
        except ValueError:  # pragma: no cover
            return None  # pragma: no cover

        can_use_simple_search = not self.is_regex and not multiline
        if can_use_simple_search and self.search_string.lower() != self.search_string.upper():
            # casing might make a difference
            can_use_simple_search = self.match_case

        return SearchData(
            regex=regex,
            word_separators=get_map_for_word_separators(self.word_separators, []) if self.word_separators else None,
            simple_search=self.search_string if can_use_simple_search else None,
        )

    def is_multiline_regex_source(self) -> bool:
        """
        Determine if the regex pattern spans multiple lines based on the provided logic.
        """
        if not self.search_string:
            return False  # pragma: no cover

        i = 0
        len_str = len(self.search_string)
        while i < len_str:
            ch_code = ord(self.search_string[i])
            if ch_code == 10:  # LineFeed
                return True
            if ch_code == 92:  # Backslash
                i += 1
                if i >= len_str:
                    break
                next_ch_code = ord(self.search_string[i])
                if next_ch_code in {110, 114, 87, 115}:  # 'n', 'r', 'W', 's'
                    return True
            i += 1
        return False

    def create_regexp(self) -> re.Pattern:
        """
        Creates a regex object from the search string and options provided.
        """
        if not self.search_string:
            raise ValueError("Cannot create regex from empty string")  # pragma: no cover

        search_string = self.search_string
        if not self.is_regex:
            search_string = escape_regexp_characters(search_string)

        flags = 0
        if not self.match_case:
            flags |= re.IGNORECASE
        # Try to create a regex out of the params
        multiline: bool
        if self.is_regex:
            multiline = self.is_multiline_regex_source()
        else:
            multiline = "\n" in self.search_string
        if multiline:
            flags |= re.MULTILINE
        flags |= re.UNICODE

        return re.compile(search_string, flags)


class LineFeedCounter(BaseModel):
    """
    LineFeedCounter is a helper class to count line feeds in a given text.

    Usage example:
        lfc = LineFeedCounter(text="your text here")
        count = lfc.find_line_feed_count_before_offset(10)
    """

    text: str
    _line_feeds_offsets: list[int]

    def __init__(self, **data: Any) -> None:
        super().__init__(**data)
        self._line_feeds_offsets = [i for i, char in enumerate(self.text) if char == "\n"]

    def find_line_feed_count_before_offset(self, offset: int) -> int:
        """
        Find the number of line feeds before the given offset using binary search.
        Note that the offset is zero-based.

        找到给定偏移量之前的换行符数量，使用二分查找。
        注意这里的Offset是从零开始的
        """
        line_feeds_offsets = self._line_feeds_offsets
        low, high = 0, len(line_feeds_offsets) - 1

        if high == -1 or offset <= line_feeds_offsets[0]:
            return 0

        while low < high:
            mid = (low + high) // 2
            if line_feeds_offsets[mid] >= offset:
                high = mid - 1
            else:
                if mid + 1 <= high and line_feeds_offsets[mid + 1] >= offset:
                    return mid + 1
                low = mid + 1

        return low + 1 if line_feeds_offsets[low] < offset else low


class TextModelSearch(BaseModel):
    @staticmethod
    def find_matches(
        model: TextModel,
        search_params: SearchParams,
        search_range: Range,
        capture_matches: bool,
        limit_result_count: int | None = None,
    ) -> list[SearchResult]:
        """
        Finds matches in a given model based on the specified search parameters and range.

        Args:
            model: The TextModel object to search in.
            search_params: The SearchParams object that contains the search criteria.
            search_range: The optional Range object defining the search range. If None, the entire model will be
                searched.
            capture_matches: A boolean value indicating whether to capture the matching text or not.
            limit_result_count: An optional integer value specifying the maximum number of results to return. If None,
                all matches will be returned.

        Returns:
            list[Range]: A list of match objects containing the range of the match and the matching text if
                capture_matches is True.
        """
        searchData = search_params.parse_search_request()
        if not searchData:
            return []  # pragma: no cover

        if searchData.regex.flags & re.MULTILINE:
            return TextModelSearch._do_find_matches_multiline(
                model,
                search_range,
                searchData.regex,
                capture_matches,
                limit_result_count,
            )
        return TextModelSearch._do_find_matches_line_by_line(
            model,
            search_range,
            searchData,
            capture_matches,
            limit_result_count,
        )

    @staticmethod
    def _do_find_matches_multiline(
        model: TextModel,
        search_range: Range,
        regex: re.Pattern,
        capture_matches: bool,
        limit_result_count: int | None,
    ) -> list[SearchResult]:
        # Get the start position's delta offset within the entire model content
        delta_offset = model.get_offset_at(search_range.start_position)
        # Retrieve the text in the specified range, using LF as the EOL preference for consistent handling
        text = model.get_value_in_range(search_range, EndOfLinePreference.LF)  # EndOfLinePreference.LF

        # Check if the model uses CRLF and instantiate LineFeedCounter if it does
        lf_counter = LineFeedCounter(text=text) if model.get_eol() == "\r\n" else None

        # Initialize result list
        result: list[SearchResult] = []

        # Iterate over regex matches in the text
        matches = list(regex.finditer(text))
        for match in matches:
            if len(result) >= (limit_result_count or float("inf")):  # Handle None as no limit
                break  # pragma: no cover
            # Use the modified _get_multiline_match_range to calculate the match range
            match_range = TextModelSearch._get_multiline_match_range(
                model,
                delta_offset,
                lf_counter,
                match.start(),
                match.group(),
            )
            # Append the match details to the result list
            result.append(SearchResult(range=match_range, match=match.group() if capture_matches else None))

        return result

    @staticmethod
    def _do_find_matches_line_by_line(
        model: TextModel,
        search_range: Range,
        searchData: SearchData,
        capture_matches: bool,
        limit_result_count: int | None,
    ) -> list[SearchResult]:
        result: list[SearchResult] = []
        # Loop through each line in the range and apply the regex
        for line_number in range(search_range.start_position.line, search_range.end_position.line + 1):
            line_text = model.get_line_content(line_number)
            matches = list(searchData.regex.finditer(line_text))

            for match in matches:
                if limit_result_count is not None and len(result) >= limit_result_count:
                    break  # pragma: no cover
                result.append(
                    SearchResult(
                        range=Range(
                            start_position=Position(
                                line_number,
                                match.start() + 1,
                            ),  # IDE中的Position从1开始计数，所以+1。 Position is 1-based in IDEs, so +1.
                            end_position=Position(
                                line_number,
                                match.end() + 1,
                            ),  # IDE中的Position从1开始计数，所以+1。 Position is 1-based in IDEs, so +1.
                        ),
                        match=match.group() if capture_matches else None,
                    ),
                )

        return result

    @staticmethod
    def _get_multiline_match_range(
        model: TextModel,
        delta_offset: int,
        lf_counter: LineFeedCounter | None,
        match_index: int,
        match0: str,
    ) -> Range:
        """
        Adjust the match range for different EOL conventions, accounting for CRLF where needed.

        Args:
            model: The TextModel object to search in.
            delta_offset: The offset of the start position of the search range.
            lf_counter: The LineFeedCounter object to count line feeds.
            match_index: The index of the match in the text.
            match0: The text of the match.

        Returns:
            Range: The adjusted range of the match.
        """
        if lf_counter:
            line_feed_count_before_match = lf_counter.find_line_feed_count_before_offset(match_index)
            start_offset = delta_offset + match_index + line_feed_count_before_match
            line_feed_count_before_end_of_match = lf_counter.find_line_feed_count_before_offset(
                match_index + len(match0),
            )
            line_feed_count_in_match = line_feed_count_before_end_of_match - line_feed_count_before_match
            end_offset = start_offset + len(match0) + line_feed_count_in_match
        else:
            start_offset = delta_offset + match_index
            end_offset = start_offset + len(match0)

        start_position = model.get_position_at(start_offset)
        end_position = model.get_position_at(end_offset)
        return Range(
            start_position=Position(start_position[0], start_position[1]),
            end_position=Position(end_position[0], end_position[1]),
        )
