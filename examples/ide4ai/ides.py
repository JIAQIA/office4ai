# filename: ides.py
# @Time    : 2024/5/13 15:12
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import threading
from typing import Any

from ide4ai.base import WorkspaceSetting
from ide4ai.environment.terminal.command_filter import CommandFilterConfig
from ide4ai.python_ide.ide import PythonIDE


class IDESingleton(type):
    _instances: dict = {}
    _lock = threading.Lock()

    def __call__(cls, *args: Any, **kwargs: Any) -> Any:
        # Create a key by combining the class name with arguments and keyword arguments
        key = cls.__name__ + kwargs.get("project_name", "")
        if key not in cls._instances:
            with cls._lock:
                if key not in cls._instances:
                    # Use super to call the parent class's __call__ method, which ultimately invokes the target class's
                    # constructor
                    cls._instances[key] = super().__call__(*args, **kwargs)
        return cls._instances[key]


class PyIDESingleton(metaclass=IDESingleton):
    """
    A thread-safe singleton class for storing the PyIDE instance.
    """

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
        self._ide: PythonIDE = PythonIDE(
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

    @property
    def ide(self) -> PythonIDE:
        return self._ide
