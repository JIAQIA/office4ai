# filename: pexpect_terminal_env.py
# @Time    : 2025/10/28 18:26
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
基于 pexpect 的终端环境实现 | Terminal environment implementation based on pexpect

相比于 local_terminal_env.py 的优势:
1. 持久会话 - 维持一个持久的 shell 进程,保持环境变量和状态
2. 虚拟环境支持 - 可以在初始化时激活虚拟环境,后续命令都在该环境中执行
3. 交互式支持 - 支持需要用户输入的命令
4. 真实 shell 行为 - cd、export 等内置命令正常工作
"""

import os
import re
import time
from typing import Any, ClassVar, cast

import gymnasium as gym
import pexpect
from loguru import logger
from typing_extensions import SupportsFloat

from ide4ai.environment.terminal.base import BaseTerminalEnv, EnvironmentArguments
from ide4ai.environment.terminal.command_filter import CommandFilterConfig
from ide4ai.schema import IDEAction, IDEObs


class PexpectTerminalEnv(BaseTerminalEnv):
    """
    基于 pexpect 的终端环境 | Terminal environment based on pexpect

    使用持久的 shell 会话来执行命令,支持虚拟环境激活和状态保持
    Uses persistent shell session to execute commands, supports virtual environment activation and state persistence

    Attributes:
        name (str): 环境名称 | Environment name
        cmd_filter (CommandFilterConfig): 命令过滤配置(黑白名单) | Command filter configuration (blacklist/whitelist)
        work_dir (str): 工作目录 | Working directory
        active_venv_cmd (str | None): 虚拟环境初始化命令 | Virtual environment initialization command
        shell (pexpect.spawn): 持久的 shell 进程 | Persistent shell process
    """

    name: ClassVar[str] = "PexpectTerminalEnv"
    metadata: dict[str, Any] = {"render_modes": ["ansi"]}

    # Shell 提示符模式 | Shell prompt pattern
    PROMPT_PATTERN = r"[\$#>]\s*$"

    def __init__(
        self,
        args: EnvironmentArguments,
        work_dir: str,
        cmd_filter: CommandFilterConfig | None = None,
        active_venv_cmd: str | None = None,
        shell: str = "/bin/bash",
    ) -> None:
        """
        初始化 pexpect 终端环境 | Initialize pexpect terminal environment

        Args:
            args: 环境参数 | Environment arguments
            work_dir: 工作目录 | Working directory
            cmd_filter: 命令过滤配置(黑白名单) | Command filter configuration (blacklist/whitelist)
            active_venv_cmd: 虚拟环境初始化命令,例如 "source .venv/bin/activate" 或 "uv venv activate"
                      Virtual environment initialization command, e.g., "source .venv/bin/activate" or "uv venv activate"
            shell: Shell 程序路径 | Shell program path
        """
        super().__init__()
        self.args = args

        # 处理命令过滤配置 | Handle command filter config
        if cmd_filter is not None:
            self.cmd_filter = cmd_filter
        else:
            # 默认使用黑名单模式 | Default to blacklist mode
            self.cmd_filter = CommandFilterConfig.allow_all_except()

        self.active_venv_cmd = active_venv_cmd
        self.shell_path = shell

        # 验证工作目录 | Validate working directory
        if os.path.exists(work_dir) and os.path.isdir(work_dir):
            self.work_dir = self.current_dir = os.path.expanduser(work_dir)
        else:
            raise ValueError(f"Work directory {work_dir} does not exist")

        # 状态标志 | State flags
        self._is_closing = False
        self._is_closed = False
        # 超时设置 | Timeout settings
        self.timeout = self.args.timeout

        # 命令历史记录 | Command history
        self._command_history: list[dict[str, str]] = []

        # 虚拟环境激活状态 | Virtual environment activation status
        self.venv_activated = False
        self.venv_activation_error: str | None = None

        # 初始化持久 shell 会话 | Initialize persistent shell session
        self._init_shell()

        # Gym spaces
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

    def _init_shell(self) -> None:
        """
        初始化持久的 shell 会话 | Initialize persistent shell session

        启动一个 shell 进程并进行必要的配置:
        1. 在工作目录中启动 shell
        2. 设置 PS1 提示符以便识别命令完成
        3. 如果指定了虚拟环境,激活它

        Start a shell process and perform necessary configuration:
        1. Start shell in working directory
        2. Set PS1 prompt for command completion detection
        3. Activate virtual environment if specified
        """
        try:
            # 启动 shell 进程,直接在工作目录中启动 | Start shell process directly in working directory
            self.shell = pexpect.spawn(
                self.shell_path,
                encoding="utf-8",
                echo=False,
                timeout=self.timeout,
                cwd=self.work_dir,
            )

            # 设置简单的提示符以便于匹配 | Set simple prompt for easy matching
            self.shell.sendline('export PS1="PEXPECT_PROMPT> "')
            self.shell.expect("PEXPECT_PROMPT>", timeout=5)

            # 激活虚拟环境(如果指定) | Activate virtual environment (if specified)
            if self.active_venv_cmd:
                try:
                    self.shell.sendline(self.active_venv_cmd)
                    index = self.shell.expect(
                        ["PEXPECT_PROMPT>", pexpect.TIMEOUT, pexpect.EOF],
                        timeout=10,
                    )

                    if index == 0:
                        # 检查命令退出码 | Check command exit code
                        self.shell.sendline("echo $?")
                        self.shell.expect("PEXPECT_PROMPT>", timeout=5)
                        exit_code_output = self.shell.before or ""

                        # 提取退出码 | Extract exit code
                        exit_code_match = re.search(r"(\d+)", exit_code_output)
                        exit_code = int(exit_code_match.group(1)) if exit_code_match else 1

                        if exit_code == 0:
                            # 退出码为0,激活成功 | Exit code is 0, activation successful
                            self.venv_activated = True
                            logger.info(f"虚拟环境激活成功 | Virtual environment activated: {self.active_venv_cmd}")
                        else:
                            # 退出码非0,激活失败 | Exit code is non-zero, activation failed
                            self.venv_activated = False
                            self.venv_activation_error = (
                                f"虚拟环境激活命令返回非零退出码: {exit_code} | "
                                f"Venv activation command returned non-zero exit code: {exit_code}"
                            )
                            logger.warning(
                                f"虚拟环境激活失败: {self.venv_activation_error} | "
                                f"Virtual environment activation failed: {self.venv_activation_error}",
                            )
                    else:
                        # 激活超时或失败 | Activation timeout or failed
                        self.venv_activated = False
                        self.venv_activation_error = "虚拟环境激活超时 | Virtual environment activation timeout"
                        logger.warning(
                            f"虚拟环境激活失败: {self.venv_activation_error} | "
                            f"Virtual environment activation failed: {self.venv_activation_error}",
                        )

                except (pexpect.TIMEOUT, pexpect.EOF) as venv_error:
                    # 虚拟环境激活失败,但不影响 shell 初始化 | Venv activation failed, but don't fail shell init
                    self.venv_activated = False
                    self.venv_activation_error = str(venv_error)
                    logger.warning(
                        f"虚拟环境激活失败: {venv_error} | Virtual environment activation failed: {venv_error}",
                    )

        except (pexpect.TIMEOUT, pexpect.EOF) as e:
            raise RuntimeError(f"Failed to initialize shell: {e}") from e

    def construct_action(self, action: dict) -> IDEAction:
        """
        构建 IDEAction 对象 | Construct IDEAction object

        Args:
            action: 动作字典 | Action dictionary

        Returns:
            IDEAction 对象 | IDEAction object

        Raises:
            ValueError: 如果动作不合法 | If action is invalid
        """
        ide_action = IDEAction.model_validate(action)

        if ide_action.category != "terminal":
            raise ValueError(f"Unsupported action category: {ide_action.category}")

        if not self.cmd_filter.is_allowed(ide_action.action_name):
            reason = self.cmd_filter.get_rejection_reason(ide_action.action_name)
            raise ValueError(reason)

        if not isinstance(ide_action.action_args, (list, str)):
            raise ValueError(
                f"Unsupported action arguments: {ide_action.action_args}, args should be str or list[str]",
            )

        return ide_action

    def step(self, action: dict) -> tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]:
        """
        执行一个动作 | Execute an action

        Args:
            action: 动作字典,包含 category, action_name, action_args
                   Action dictionary containing category, action_name, action_args

        Returns:
            观察、奖励、是否结束、是否成功、额外信息
            Observation, reward, done, success, extra info
        """
        self._assert_not_closed()

        # 构建并验证动作 | Construct and validate action
        terminal_action = self.construct_action(action)
        cmd = terminal_action.action_name
        args = (
            [terminal_action.action_args]
            if isinstance(terminal_action.action_args, str)
            else terminal_action.action_args
        )

        # 构建完整命令 | Build complete command
        full_command = f"{cmd} {' '.join(cast(list[str], args))}" if args else cmd

        # 执行命令 | Execute command
        output, success = self._execute_command(full_command)

        # 记录命令历史 | Record command history
        self._command_history.append(
            {
                "command": full_command,
                "output": output,
                "success": str(success),
            },
        )

        # 返回观察结果 | Return observation
        obs = IDEObs(obs=output)
        reward = 100.0 if success else 0.0
        done = True  # 命令执行完成 | Command execution completed

        return obs.model_dump(), reward, done, success, {}

    def _execute_command(self, command: str) -> tuple[str, bool]:
        """
        在持久 shell 会话中执行命令 | Execute command in persistent shell session

        Args:
            command: 要执行的命令 | Command to execute

        Returns:
            命令输出和是否成功 | Command output and success status
        """
        try:
            # 发送命令 | Send command
            self.shell.sendline(command)

            # 等待命令完成 | Wait for command completion
            index = self.shell.expect(
                ["PEXPECT_PROMPT>", pexpect.TIMEOUT, pexpect.EOF],
                timeout=self.timeout,
            )

            if index == 0:
                # 命令正常完成 | Command completed normally
                output = self.shell.before

                # 检查退出状态 | Check exit status
                self.shell.sendline("echo $?")
                self.shell.expect("PEXPECT_PROMPT>", timeout=5)
                exit_code_output = self.shell.before

                # 提取退出码 | Extract exit code
                exit_code_match = re.search(r"(\d+)", exit_code_output or "")
                exit_code = int(exit_code_match.group(1)) if exit_code_match else 1

                success = exit_code == 0

                # 清理输出 | Clean output
                output = self._clean_output(output or "")

                return output, success

            elif index == 1:
                # 超时 | Timeout
                return f"Command timeout after {self.timeout} seconds", False

            else:
                # EOF - shell 进程意外终止 | EOF - shell process terminated unexpectedly
                return "Shell process terminated unexpectedly", False

        except Exception as e:
            return f"Error executing command: {str(e)}", False

    @staticmethod
    def _clean_output(output: str) -> str:
        """
        清理命令输出,移除控制字符和多余空白 | Clean command output, remove control characters and extra whitespace

        Args:
            output: 原始输出 | Raw output

        Returns:
            清理后的输出 | Cleaned output
        """
        # 移除 ANSI 转义序列 | Remove ANSI escape sequences
        ansi_escape = re.compile(r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])")
        output = ansi_escape.sub("", output)

        # 移除回车符 | Remove carriage returns
        output = output.replace("\r", "")

        # 移除首尾空白 | Strip leading/trailing whitespace
        output = output.strip()

        return output

    def reset(
        self,
        *,
        seed: int | None = None,
        options: dict[str, Any] | None = None,
    ) -> tuple[IDEObs, dict[str, Any]]:
        """
        重置环境 | Reset environment

        重新初始化 shell 会话,清除命令历史
        Reinitialize shell session, clear command history

        Args:
            seed: 随机种子 | Random seed
            options: 额外选项 | Additional options

        Returns:
            观察结果和额外信息 | Observation and extra info
        """
        self._assert_not_closed()

        # 关闭现有 shell | Close existing shell
        if hasattr(self, "shell") and self.shell.isalive():
            self.shell.close()

        # 清空命令历史 | Clear command history
        self._command_history.clear()

        # 重新初始化 shell | Reinitialize shell
        self._init_shell()

        return IDEObs(obs="Reset environment successfully"), {}

    def render(self) -> str:  # type: ignore[override]
        """
        渲染当前环境状态 | Render current environment state

        返回最近的命令历史
        Return recent command history

        Returns:
            渲染结果 | Render result
        """
        self._assert_not_closed()

        if not self._command_history:
            return f"PEXPECT_PROMPT> (cwd: {self.current_dir})"

        # 渲染最近 3 条命令 | Render last 3 commands
        render_frames = []
        for entry in self._command_history[-3:]:
            render_frames.append(
                f"PEXPECT_PROMPT> {entry['command']}\n{entry['output']}",
            )

        return "\n\n".join(render_frames)

    def close(self) -> None:
        """
        关闭环境 | Close environment

        终止 shell 进程并清理资源
        Terminate shell process and clean up resources
        """
        if self._is_closed or self._is_closing:
            return

        self._is_closing = True

        try:
            if hasattr(self, "shell") and self.shell.isalive():
                # 尝试优雅退出 | Try graceful exit
                self.shell.sendline("exit")
                time.sleep(0.5)

                # 如果还活着,强制终止 | Force terminate if still alive
                if self.shell.isalive():
                    self.shell.close(force=True)
        except Exception as e:
            # 忽略关闭时的错误 | Ignore errors during close
            logger.error(f"关闭PexpectTerminal时发生异常: {e}")
        finally:
            self._command_history.clear()
            self._is_closed = True
            self._is_closing = False

    def _assert_not_closed(self) -> bool:
        """
        断言环境未关闭 | Assert environment is not closed

        Returns:
            如果环境未关闭返回 True | True if environment is not closed

        Raises:
            ValueError: 如果环境已关闭 | If environment is closed
        """
        if self._is_closed:
            raise ValueError("Environment is closed.")
        return True

    def change_dir(self, *, path: str) -> tuple[str, bool]:
        """
        更改当前目录 | Change current directory

        使用 cd 命令在持久 shell 中切换目录
        Use cd command to change directory in persistent shell

        Args:
            path: 目标目录路径 | Target directory path

        Returns:
            输出信息和是否成功 | Output message and success status
        """
        self._assert_not_closed()

        # 展开路径 | Expand path
        path = os.path.expanduser(path)

        # 验证路径是否在工作目录内 | Validate path is within working directory
        real_path = os.path.realpath(path)
        real_work_dir = os.path.realpath(self.work_dir)

        try:
            common_path = os.path.commonpath([real_path, real_work_dir])
            if common_path != real_work_dir:
                return f"Path {path} is not a subdirectory of {self.work_dir}", False
        except ValueError:
            return f"Path {path} is not a subdirectory of {self.work_dir}", False

        # 执行 cd 命令 | Execute cd command
        output, success = self._execute_command(f'cd "{path}"')

        if success:
            self.current_dir = real_path
            return f"Changed directory to {path}", True
        else:
            return f"Failed to change directory: {output}", False

    def get_env_var(self, var_name: str) -> str | None:
        """
        获取环境变量 | Get environment variable

        Args:
            var_name: 环境变量名 | Environment variable name

        Returns:
            环境变量值,如果不存在返回 None | Environment variable value, None if not exists
        """
        self._assert_not_closed()

        output, success = self._execute_command(f'echo "${var_name}"')

        if success and output:
            return output.strip()
        return None

    def set_env_var(self, var_name: str, value: str) -> bool:
        """
        设置环境变量 | Set environment variable

        Args:
            var_name: 环境变量名 | Environment variable name
            value: 环境变量值 | Environment variable value

        Returns:
            是否成功 | Success status
        """
        self._assert_not_closed()

        _, success = self._execute_command(f'export {var_name}="{value}"')
        return success
