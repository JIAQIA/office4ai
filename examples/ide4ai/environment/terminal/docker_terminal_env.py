# filename: docker_terminal_env.py
# @Time    : 2024/4/18 10:45
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
# from typing import Any, ClassVar, SupportsFloat
#
# from gymnasium.core import ActType, ObsType
#
# from ide4ai.environment.terminal.base import (
#     BaseTerminalEnv,
#     EnvironmentArguments,
# )
#
#
# class DockerTerminalEnv(BaseTerminalEnv):
#     """
#     DockerTerminalEnv is a gym environment for docker terminal.
#
#     Attributes:
#         name (str): The name of the environment.
#     """
#
#     name: ClassVar[str] = "DockerTerminalEnv"
#     metadata: dict[str, Any] = {"render_modes": ["ansi"]}
#
#     def __init__(self, args: EnvironmentArguments):
#         super(DockerTerminalEnv, self).__init__()
#         self.args = args
#         # Set timeout
#         self.timeout = self.args.timeout
#         # TODO idx如何有更好的定义
#         self.idx = 0
#
#     def step(self, action: ActType) -> tuple[ObsType, SupportsFloat, bool, bool, dict[str, Any]]:
#         pass
#
#     def reset(
#         self,
#         *,
#         seed: int | None = None,
#         options: dict[str, Any] | None = None,
#     ) -> tuple[ObsType, dict[str, Any]]:
#         pass
#
#     def render(self):
#         pass
#
#     def close(self) -> None:
#         pass
