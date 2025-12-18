# filename: ide.py
# @Time    : 2024/4/16 18:35
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from typing import Any, ClassVar

from ide4ai.base import IDE, WorkspaceSetting
from ide4ai.environment.terminal.base import EnvironmentArguments
from ide4ai.environment.terminal.command_filter import CommandFilterConfig
from ide4ai.environment.terminal.pexpect_terminal_env import PexpectTerminalEnv
from ide4ai.python_ide.workspace import PyWorkspace


class PythonIDE(IDE[PexpectTerminalEnv, PyWorkspace]):
    """
    PythonIDE is a gym environment for python ide.

    实现原理是基于Docker容器，在一个容器内跑一个Python版本的slim镜像，然后在这个容器内运行LSP服务，通过LSP服务来实现IDE的功能。
    而PythonIDE当前这个类的封装，在于将通用的能力通过step调用传入到容器内，然后容器内的LSP服务来处理这些能力。

    Attributes:
        name (str): The name of the environment.
        metadata (dict[str, Any]): The metadata of the environment.
        root_dir (str): The root directory of the environment.
        project_name (str): The project name of the environment.
        render_with_symbols (bool): Whether render with symbols.
        max_active_models (int): The max active models of the environment.
        enable_simple_view_mode (bool): 是否启用简化视图模式 / Whether to enable simple view mode.
    """

    name: ClassVar[str] = "PythonIDE"
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
        active_venv_cmd: str | None = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(
            root_dir=root_dir,
            project_name=project_name,
            cmd_filter=cmd_filter,
            render_with_symbols=render_with_symbols,
            max_active_models=max_active_models,
            cmd_time_out=cmd_time_out,
            enable_simple_view_mode=enable_simple_view_mode,
            workspace_setting=workspace_setting,
            **kwargs,
        )
        self.active_venv_cmd = active_venv_cmd
        self.workspace = PyWorkspace(
            self.root_dir,
            self.project_name,
            self.render_with_symbols,
            self.max_active_models,
            self.enable_simple_view_mode,
            **self._workspace_setting,
        )

    def init_terminal(self) -> PexpectTerminalEnv:
        """
        初始化终端环境 | Initialize terminal environment

        使用 PexpectTerminalEnv 以支持持久会话和虚拟环境激活
        Use PexpectTerminalEnv to support persistent session and virtual environment activation

        Returns:
            PexpectTerminalEnv: PexpectTerminalEnv 对象 | PexpectTerminalEnv object
        """

        return PexpectTerminalEnv(
            args=EnvironmentArguments(image_name="local", timeout=self.cmd_time_out),
            work_dir=self.root_dir,
            cmd_filter=self.cmd_filter,
            active_venv_cmd=self.active_venv_cmd,
        )
