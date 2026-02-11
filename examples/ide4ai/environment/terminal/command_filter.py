# filename: command_filter.py
# @Time    : 2025/11/01 19:21
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
命令过滤配置 | Command filter configuration

提供黑白名单机制来控制可执行的命令
Provides blacklist and whitelist mechanism to control executable commands
"""

from dataclasses import dataclass, field

# 默认危险命令黑名单 | Default dangerous command blacklist
DEFAULT_BLACK_LIST = [
    "rm",  # 删除文件 | Delete files
    "rmdir",  # 删除目录 | Delete directories
    "dd",  # 磁盘操作 | Disk operations
    "mkfs",  # 格式化文件系统 | Format filesystem
    "format",  # 格式化 | Format
    "fdisk",  # 磁盘分区 | Disk partition
    "parted",  # 磁盘分区 | Disk partition
    "shutdown",  # 关机 | Shutdown
    "reboot",  # 重启 | Reboot
    "halt",  # 停机 | Halt
    "poweroff",  # 关机 | Power off
    "init",  # 系统初始化 | System init
    "telinit",  # 运行级别切换 | Run level switch
]


@dataclass
class CommandFilterConfig:
    """
    命令过滤配置 | Command filter configuration

    支持黑白名单两种模式:
    1. 如果指定了白名单(white_list不为None),必须命中才放行
    2. 如果未指定白名单(white_list为None),则仅检查黑名单

    Supports two modes:
    1. If whitelist is specified (white_list is not None), only whitelisted commands are allowed
    2. If whitelist is not specified (white_list is None), only blacklist is checked

    Attributes:
        white_list: 命令白名单,None表示不启用白名单模式 | Command whitelist, None means whitelist mode is disabled
        black_list: 命令黑名单,None表示不启用黑名单 | Command blacklist, None means blacklist is disabled
    """

    white_list: list[str] | None = None
    black_list: list[str] | None = field(default_factory=lambda: DEFAULT_BLACK_LIST.copy())

    def is_allowed(self, command: str) -> bool:
        """
        判断命令是否允许执行 | Check if command is allowed to execute

        Args:
            command: 要检查的命令 | Command to check

        Returns:
            是否允许执行 | Whether execution is allowed
        """
        # 如果指定了白名单,必须命中才放行 | If whitelist is specified, must match to allow
        if self.white_list is not None:
            return command in self.white_list

        # 如果未指定白名单,则仅检查黑名单 | If whitelist not specified, only check blacklist
        if self.black_list is not None:
            return command not in self.black_list

        # 都未指定,默认允许 | If neither specified, allow by default
        return True

    def get_rejection_reason(self, command: str) -> str:
        """
        获取命令被拒绝的原因 | Get reason why command was rejected

        Args:
            command: 被拒绝的命令 | Rejected command

        Returns:
            拒绝原因 | Rejection reason
        """
        if self.white_list is not None and command not in self.white_list:
            return f"Command '{command}' not in whitelist"

        if self.black_list is not None and command in self.black_list:
            return f"Command '{command}' is in blacklist (dangerous command)"

        return f"Command '{command}' is not allowed"

    @classmethod
    def from_white_list(cls, white_list: list[str]) -> "CommandFilterConfig":
        """
        从白名单创建配置(向后兼容) | Create config from whitelist (backward compatible)

        Args:
            white_list: 命令白名单 | Command whitelist

        Returns:
            CommandFilterConfig 实例 | CommandFilterConfig instance
        """
        return cls(white_list=white_list, black_list=None)

    @classmethod
    def allow_all_except(cls, black_list: list[str] | None = None) -> "CommandFilterConfig":
        """
        创建仅使用黑名单的配置 | Create config with only blacklist

        Args:
            black_list: 命令黑名单,None则使用默认黑名单 | Command blacklist, None to use default

        Returns:
            CommandFilterConfig 实例 | CommandFilterConfig instance
        """
        if black_list is None:
            black_list = DEFAULT_BLACK_LIST.copy()
        return cls(white_list=None, black_list=black_list)

    @classmethod
    def allow_all(cls) -> "CommandFilterConfig":
        """
        创建允许所有命令的配置(不推荐用于生产环境) | Create config allowing all commands (not recommended for production)

        Returns:
            CommandFilterConfig 实例 | CommandFilterConfig instance
        """
        return cls(white_list=None, black_list=None)
