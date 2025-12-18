# filename: terminal_env.py
# @Time    : 2024/4/16 15:10
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, ClassVar

import gymnasium as gym
from simple_parsing.helpers import FrozenSerializable
from typing_extensions import SupportsFloat

from ide4ai.schema import IDEObs


@dataclass(frozen=True)
class EnvironmentArguments(FrozenSerializable):
    """
    Configure data sources and setup instructions for th environment in which we solve the tasks.
    """

    image_name: str
    split: str = "dev"
    container_name: str | None = None
    install_environment: bool = True
    timeout: int = 35
    verbose: bool = False
    no_mirror: bool = False


class BaseTerminalEnv(gym.Env, ABC):
    """
    BaseTerminalEnv is a gym environment for terminal.

    Attributes:
        name (str): The name of the terminal environment.
    """

    name: ClassVar[str]
    metadata: dict[str, Any] = {"render_modes": ["ansi"]}

    @abstractmethod
    def step(self, action: dict) -> tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]:
        pass

    @abstractmethod
    def reset(
        self,
        *,
        seed: int | None = None,
        options: dict[str, Any] | None = None,
    ) -> tuple[IDEObs, dict[str, Any]]:
        pass

    @abstractmethod
    def render(self) -> str:  # type: ignore
        pass

    @abstractmethod
    def close(self) -> None:
        pass
