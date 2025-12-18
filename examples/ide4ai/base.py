# filename: base.py
# @Time    : 2024/4/16 19:56
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import atexit
import subprocess
from abc import ABC, abstractmethod
from collections.abc import Callable
from typing import Any, ClassVar, Generic, SupportsFloat, TypeVar

import gymnasium as gym
from typing_extensions import TypedDict

from ide4ai.environment.terminal.base import BaseTerminalEnv
from ide4ai.environment.terminal.command_filter import CommandFilterConfig
from ide4ai.environment.workspace.base import BaseWorkspace
from ide4ai.exceptions import IDEExecutionError
from ide4ai.schema import IDEAction, IDEObs

# 定义泛型类型变量用于 Terminal 和 Workspace 类型 | Define generic type variables for Terminal and Workspace types
TerminalT = TypeVar("TerminalT", bound=BaseTerminalEnv)
WorkspaceT = TypeVar("WorkspaceT", bound=BaseWorkspace)


class WorkspaceSetting(TypedDict, total=False):
    """
    工作区配置项 / Workspace configuration options

    Attributes:
        header_generators (dict[str, Callable[[BaseWorkspace, str], str]] | None):
            文件头生成器字典，key为文件扩展名（如".py"），value为生成器函数 /
            File header generator dictionary, key is file extension (e.g., ".py"), value is generator function
        shortcut_commands (dict[str, list[str]] | None):
            项目快捷命令字典，key为命令前缀（如"make"、"poe"），value为可用命令列表 /
            Project shortcut commands dictionary, key is command prefix (e.g., "make", "poe"), value is available commands list
    """

    header_generators: dict[str, Callable[[BaseWorkspace, str], str]] | None
    shortcut_commands: dict[str, list[str]] | None


class IDE(gym.Env, ABC, Generic[TerminalT, WorkspaceT]):
    """
    实现原理是基于Docker容器，在一个容器内跑一个Python版本的slim镜像，然后在这个容器内运行LSP服务，通过LSP服务来实现IDE的功能。
    而PythonIDE当前这个类的封装，在于将通用的能力通过step调用传入到容器内，然后容器内的LSP服务来处理这些能力。

    Attributes:
        name (str): 环境名称 / The name of the environment.
        metadata (dict[str, Any]): 环境元数据 / The metadata of the environment.
        root_dir (str): 环境根目录 / The root directory of the environment.
        project_name (str): 项目名称 / The project name of the environment.
        render_with_symbols (bool): 是否使用符号渲染 / Whether render with symbols.
        max_active_models (int): 最大活跃模型数量 / The max active models of the environment.
        cmd_time_out (int): 命令超时时间（秒）/ Command timeout in seconds.
        enable_simple_view_mode (bool): 是否启用简化视图模式 / Whether to enable simple view mode.
        _workspace_setting (WorkspaceSetting | None): 工作区配置项 / Workspace configuration options.
    """

    name: ClassVar[str] = "IDE"
    metadata: dict[str, Any] = {"render_modes": ["ansi"]}

    def __init__(
        self,
        root_dir: str,
        project_name: str,
        cmd_filter: CommandFilterConfig | None = None,
        render_with_symbols: bool = True,
        max_active_models: int = 3,
        cmd_time_out: int = 10,
        enable_simple_view_mode: bool = True,
        workspace_setting: WorkspaceSetting | None = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(**kwargs)

        # 处理命令过滤配置 | Handle command filter config
        if cmd_filter is not None:
            self.cmd_filter = cmd_filter
        else:
            # 默认使用黑名单模式 | Default to blacklist mode
            self.cmd_filter = CommandFilterConfig.allow_all_except()

        self.cmd_time_out = cmd_time_out
        self.root_dir = root_dir
        self.project_name = project_name
        self.render_with_symbols = render_with_symbols
        self.enable_simple_view_mode = enable_simple_view_mode
        self.max_active_models = max_active_models
        self.terminals: list[TerminalT] = []
        self.active_terminal_index: int | None = None
        self.workspace: WorkspaceT | None = None
        self._workspace_setting = workspace_setting or {}
        # 初始化动作空间与观察空间
        self.action_space = gym.spaces.Dict(
            {
                "category": gym.spaces.Discrete(2),
                "action_name": gym.spaces.Text(100),
                "action_args": gym.spaces.Text(1000),
            },
        )
        self.observation_space = gym.spaces.Dict(
            {
                "created_at": gym.spaces.Text(100),
                "obs": gym.spaces.Text(100000),
            },
        )
        atexit.register(self.close)

    def __del__(self) -> None:
        self.close()

    @property
    def terminal(self) -> TerminalT:
        if self.active_terminal_index is not None:
            return self.terminals[self.active_terminal_index]
        else:
            terminal = self.init_terminal()
            self.terminals.append(terminal)
            self.active_terminal(len(self.terminals) - 1)
            return terminal

    @abstractmethod
    def init_terminal(self) -> TerminalT: ...

    def active_terminal(self, index: int) -> None:
        self.active_terminal_index = index

    def construct_action(self, action: dict) -> IDEAction:
        """
        构建 IDEAction 对象

        Args:
            action (dict): 动作字典 | Action dictionary

        Returns:
            IDEAction: IDEAction 对象 | IDEAction object

        Raises:
            ValueError: 如果动作不在支持的动作集合中 | If the action is not in the supported action set
        """
        ide_action = IDEAction.model_validate(action)
        if ide_action.category == "terminal" and not self.cmd_filter.is_allowed(ide_action.action_name):
            reason = self.cmd_filter.get_rejection_reason(ide_action.action_name)
            err = f"{reason}. Can't run this command now."
            raise IDEExecutionError(message=err, detail_for_llm=err)
        return ide_action

    def step(self, action: dict) -> tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]:
        """
        执行一个动作

        观察返回：
        1. OpenFile: 返回打开文件的内容
        2. ApplyEdit: 返回编辑的变更记录

        奖励机制：
        1. OpenFile: 成功打印返回100，打开失败返回0
        2. ApplyEdit: 变更成功返回100，失败返回0

        Args:
            action (dict): 动作字典 | Action dictionary

        Returns:
            tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]: 观察、奖励、是否结束、是否成功、额外信息 |
                Observation, Reward, Done, Success, Extra info

        Raises:
            ValueError: 如果工作区尚未正常初始化 | If the workspace has not been initialized properly
        """
        ide_action = self.construct_action(action)
        if ide_action.category == "terminal":
            return self.terminal.step(action)
        else:
            if self.workspace:
                return self.workspace.step(action)
            else:
                raise IDEExecutionError(
                    "Workspace is not initialized",
                    detail_for_llm="Workspace is not initialized, initialize workspace first",
                )

    def reset(
        self,
        *,
        seed: int | None = None,
        options: dict[str, Any] | None = None,
    ) -> tuple[IDEObs, dict[str, Any]]:
        """
        Resets the environment.

        This function resets the environment to its initial state, including resetting the workspace and all terminals.

        Args:
            seed (int | None): The seed to be used for resetting the environment. Defaults to None.
            options (dict[str, Any] | None): Additional options for the reset operation. Defaults to None.

        Returns:
            tuple[IDEObs, dict[str, Any]]: A tuple containing the observation result and additional information.
        """
        super().reset(seed=seed, options=options)
        # 重置工作区与所有终端 | Reset workspace and all terminals
        if self.workspace:
            self.workspace.reset(seed=seed, options=options)
        if self.terminals:
            for terminal in self.terminals:
                terminal.reset(seed=seed, options=options)
        return IDEObs(obs="Reset IDE successfully"), {}

    def _get_git_status(self) -> str | None:
        """
        获取 git 仓库状态

        使用 subprocess 独立执行 git status -s，不影响 terminal 的命令历史
        Use subprocess to execute git status -s independently, without affecting terminal command history

        Returns:
            str | None: git status -s 的输出，如果不是 git 仓库或执行失败则返回 None
                       Output of git status -s, or None if not a git repo or execution failed
        """
        try:
            # 使用 subprocess 执行 git status -s
            # 设置 cwd 为 root_dir，确保在正确的目录执行
            result = subprocess.run(
                ["git", "status", "-s"],
                cwd=self.root_dir,
                capture_output=True,
                text=True,
                timeout=5,  # 5秒超时，防止卡住
            )

            # 如果命令执行成功且有输出
            if result.returncode == 0:
                output = result.stdout.strip()
                # 只有当有实际内容时才返回（空字符串表示没有变更）
                return output if output else None

            return None

        except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
            # FileNotFoundError: git 命令不存在
            # TimeoutExpired: 执行超时
            # Exception: 其他异常（如不是 git 仓库）
            return None

    def render(self) -> str:  # type: ignore
        """
        渲染 IDE 内容

        Returns:
            str: 渲染结果 | Render result
        """
        # 添加 git 状态信息
        git_status = self._get_git_status()
        if git_status:
            content = f"\n当前项目 Git 状态 (git status -s):\n{git_status}\n"
        else:
            content = ""

        content += "IDE Content:\n"
        if self.workspace:
            content += f"当前工作区内容如下:\n{self.workspace.render()}\n"
        if self.active_terminal_index is not None:
            content += f"当前终端内容如下:\n{self.terminal.render()}\n"

        return content

    def close(self) -> None:
        for t in self.terminals:
            t.close()
        if self.workspace:
            self.workspace.close()
