# filename: character_classifier.py
# @Time    : 2024/5/6 17:11
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm


class CharacterClassifier:
    """
    Initialization:
        Initialize the ascii_map as a byte array (bytearray) for ASCII character codes (0-255), filled with the default
            value.
        A dictionary (self.map) handles non-ASCII character codes.
    Type Conversion:
        The to_uint8 function mimics JavaScript's Uint8 by using bitwise operations to keep values in the 0-255 range.
    Set and Get Methods:
        The set method updates the value for a given character code, using the ascii_map for ASCII values and the map
            for others.
        The get method retrieves the value for a character code, returning the default value if the code is not
            explicitly set.
    Clear Method:
        Resets the ascii_map to the default value and clears the dictionary for non-ASCII character codes.

    这是关于一个以字符代码为键的映射的初始化和操作的描述：
    初始化：
        初始化ascii_map作为一个字节数组（bytearray），用于ASCII字符代码（0-255），并填充默认值。
        使用一个字典（self.map）来处理非ASCII字符代码。
    类型转换：
        to_uint8函数模仿JavaScript的Uint8，使用位操作将值保持在0-255的范围内。
    设置与获取方法：
        set方法更新给定字符代码的值，对于ASCII值使用ascii_map，对于其他值使用map。
        get方法检索字符代码的值，如果代码未显式设置，则返回默认值。
    清除方法：
        重置ascii_map为默认值，并清除非ASCII字符代码的字典。
    """

    def __init__(self, default_value: int):
        self.default_value: int = self.to_uint8(default_value)
        self.ascii_map: bytearray = self.create_ascii_map(self.default_value)
        self.map: dict[int, int] = {}

    @staticmethod
    def to_uint8(value: int) -> int:
        """Ensure the value is within 0-255 range, similar to Uint8 in JavaScript."""
        return value & 0xFF

    @staticmethod
    def create_ascii_map(default_value: int) -> bytearray:
        """Create a byte array for ASCII values initialized to the default value."""
        return bytearray([default_value] * 256)

    def set(self, char_code: int, value: int) -> None:
        """Set the value for a specific character code."""
        value = self.to_uint8(value)
        if 0 <= char_code < 256:
            self.ascii_map[char_code] = value
        else:
            self.map[char_code] = value

    def get(self, char_code: int) -> int:
        """Retrieve the value for a specific character code, defaulting if not set."""
        if 0 <= char_code < 256:
            return self.ascii_map[char_code]
        else:
            return self.map.get(char_code, self.default_value)

    def clear(self) -> None:
        """Clear all values, resetting ASCII map and clearing the dictionary."""
        self.ascii_map = self.create_ascii_map(self.default_value)
        self.map.clear()
