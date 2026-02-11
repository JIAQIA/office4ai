# filename: word_character_classifier.py
# @Time    : 2024/5/6 18:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from typing import cast

import regex as re
from cachetools import LRUCache

from ide4ai.environment.workspace.common.core.character_classifier import (
    CharacterClassifier,
)

# 假设已经定义了 CharacterClassifier


class WordCharacterClassifier(CharacterClassifier):
    """
    在编辑器或IDE中，WordCharacterClassifier 这类工具非常关键，主要用于处理和识别文本中的单词结构和边界。具体地说，WordCharacterClassifier 可以支持以下功能：

    1. 语法高亮
    通过识别单词和单词分隔符，WordCharacterClassifier 可以帮助编辑器实现语法高亮，区分关键字、变量、字符串等。这不仅使得代码更易于阅读，也帮助开发者快速识别语法错误。

    2. 文本选择和导航
    在文本编辑中，用户常常需要进行单词级的选择和导航（例如，通过按Ctrl+左键/右键在单词间跳转）。WordCharacterClassifier 通过定义哪些字符构成单词边界，使得这些操作更加精确和符合预期。

    3. 自动补全
    编辑器和IDE的自动补全功能依赖于对正在编辑的单词的准确识别。WordCharacterClassifier 提供了这种识别能力，使得IDE能够基于当前单词提供相关的补全选项，如方法名、属性、变量等。

    4. 代码格式化
    代码格式化工具使用 WordCharacterClassifier 来确定哪些位置可以插入换行或额外的空格，以符合编程语言的格式规范。例如，避免在单词中间断行，或在关键字后自动添加空格。

    5. 搜索和替换
    高级的搜索和替换功能可能需要对完整的单词进行操作，而不是简单的字符序列。WordCharacterClassifier 可以帮助编辑器理解何处开始和结束一个单词，从而执行精确的“全词匹配”搜索。

    6. 语言处理
    在支持多语言的环境中，WordCharacterClassifier 可以根据不同语言的规则（通过intl_segmenter_locales指定）来识别单词，这对于处理像中文、日语等没有明显空格分隔的语言尤为重要。

    7. 拼写检查
    编辑器和IDE中的拼写检查工具需要知道哪些字符组成一个可检查的单元（即单词）。WordCharacterClassifier 提供了这种单词的边界定义，使拼写检查可以正确运行。
    """

    def __init__(self, word_separators: str, intl_segmenter_locales: list[str]):
        super().__init__(0)  # Assuming 0 represents 'Regular'
        self.intl_segmenter_locales = intl_segmenter_locales
        self.word_separators = word_separators
        self._segmenter = self._create_segmenter()
        for ch in word_separators:
            self.set(ord(ch), 2)  # Assuming 2 represents 'WordSeparator'
        self.set(32, 1)  # Space
        self.set(9, 1)  # Tab

    def _create_segmenter(self) -> re.Pattern | None:
        """
        请注意，这里没有使用NLTK或者spaCy等第三方库进行更精细的分词，而是使用了Python内置的re库来实现简单的分词。
        因此这里是有改进空间的

        Returns:
            Optional[re.Pattern]: A compiled regex pattern for simple word boundary detection.
        """
        if self.intl_segmenter_locales:
            pattern = r"\w+|[^\w\s]+"  # A simple word boundary regex; refine as needed
            return re.compile(pattern)
        return None

    def find_prev_word_before_or_at_offset(self, line: str, offset: int) -> str | None:
        if not self._segmenter:
            return None
        segments = self._segmenter.finditer(line)
        last_valid = None
        for match in segments:
            if match.start() > offset:
                break
            last_valid = match.group()
        return last_valid

    def find_next_word_at_or_after_offset(self, line_content: str, offset: int) -> str | None:
        if not self._segmenter:
            return None
        segments = self._segmenter.finditer(line_content)
        for match in segments:
            if match.start() >= offset:
                return match.group()
        return None  # pragma: no cover


# Setup cache for WordCharacterClassifier instances
word_classifier_cache: LRUCache = LRUCache(maxsize=10)


def get_map_for_word_separators(word_separators: str, intl_segmenter_locales: list[str]) -> WordCharacterClassifier:
    key = f"{word_separators}/{''.join(intl_segmenter_locales)}"
    if key not in word_classifier_cache:
        classifier = WordCharacterClassifier(word_separators, intl_segmenter_locales)
        word_classifier_cache[key] = classifier
    return cast(WordCharacterClassifier, word_classifier_cache[key])
