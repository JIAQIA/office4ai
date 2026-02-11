# filename: local_terminal_env.py
# @Time    : 2024/4/18 10:45
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import os
import select
import signal
import subprocess
import time
from collections.abc import Iterator
from typing import Any, ClassVar, cast

import gymnasium as gym
from typing_extensions import SupportsFloat

from ide4ai.environment.terminal.base import BaseTerminalEnv, EnvironmentArguments
from ide4ai.environment.terminal.command_filter import CommandFilterConfig
from ide4ai.schema import IDEAction, IDEObs


class TerminalEnv(BaseTerminalEnv):
    """
    TerminalEnv is a gym environment for terminal.

    Attributes:
        name (str): The name of the environment.
        cmd_filter (CommandFilterConfig): 命令过滤配置(黑白名单) | Command filter configuration (blacklist/whitelist)
        work_dir (str): The work directory of the environment. | 工作目录
        current_dir (str): The current directory of the environment. | 当前目录
    """

    name: ClassVar[str] = "TerminalEnv"
    metadata: dict[str, Any] = {"render_modes": ["ansi"]}

    def __init__(
        self,
        args: EnvironmentArguments,
        work_dir: str,
        cmd_filter: CommandFilterConfig | None = None,
    ) -> None:
        super().__init__()
        self.args = args

        # 处理命令过滤配置 | Handle command filter config
        if cmd_filter is not None:
            self.cmd_filter = cmd_filter
        else:
            # 默认使用黑名单模式 | Default to blacklist mode
            self.cmd_filter = CommandFilterConfig.allow_all_except()
        if os.path.exists(work_dir) and os.path.isdir(work_dir):
            self.work_dir = self.current_dir = os.path.expanduser(work_dir)
        else:
            raise ValueError(f"Work directory {work_dir} does not exist")
        self._is_closing = False
        self._is_closed = False
        # Set timeout
        self.timeout = self.args.timeout
        # 当前被TerminalEnv管理的进程列表
        self.procs: dict[int, subprocess.Popen] = {}
        # 用于记录每个进程的输出结果
        self._procs_result: dict[int, Iterator[tuple[str, bool]]] = {}
        # 用于记录每个进程的最终成功状态 / Record final success status for each process
        self._procs_final_success: dict[int, bool] = {}
        # 用户记录每个进程启动时指定的工作目录
        self._procs_working_description: dict[int, str] = {}
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

    def construct_action(self, action: dict) -> IDEAction:
        """
        构建 IDEAction 对象

        Args:
            action (dict): 动作字典 | Action dictionary

        Returns:
            IDEAction: IDEAction 对象 | IDEAction object

        Raises:
            ValueError: 如果动作类别为 "terminal" | If the action category is "terminal"
            ValueError: 如果动作不在支持的动作集合中 | If the action is not in the supported action set
        """
        ide_action = IDEAction.model_validate(action)
        match ide_action.category:
            case "terminal":
                if not self.cmd_filter.is_allowed(ide_action.action_name):
                    reason = self.cmd_filter.get_rejection_reason(ide_action.action_name)
                    raise ValueError(reason)
                elif not isinstance(ide_action.action_args, list) and not isinstance(ide_action.action_args, str):
                    raise ValueError(
                        f"Unsupported action arguments: {ide_action.action_args}, args should be str or list[str]",
                    )
                else:
                    return ide_action
            case _:
                raise ValueError(f"Unsupported action category: {ide_action.category}")

    def step(self, action: dict) -> tuple[dict, SupportsFloat, bool, bool, dict[str, Any]]:
        """
        执行一个动作

        Args:
            action (dict): An instance of the IDEAction class representing the action to be performed.

        Returns:
            A tuple containing the following elements:
            - An instance of the IDEObs class representing the observation after performing the action.
            - An instance of SupportsFloat representing the reward obtained after performing the action.
            - A boolean value indicating whether the current episode is done or not.
            - A boolean value indicating whether the action performed was successful or not.
            - A dictionary containing additional information about the action performed.

        """
        self._assert_not_closed()
        terminal_action = self.construct_action(action)
        cmd = terminal_action.action_name
        args = (
            [terminal_action.action_args]
            if isinstance(terminal_action.action_args, str)
            else terminal_action.action_args
        )
        pid = self.run(cmd=cmd, args=cast(list[str], args))
        obs_str, done, success = self.get_proc_res(pid=pid)
        # 如果当前未完成，则进行间隔1秒的采样，直到timeout或者完成
        total_time: int = 0
        while not done and total_time < self.timeout:
            total_time += 1
            time.sleep(1)
            yield_obs_str, done, success = self.get_proc_res(pid=pid)
            obs_str += yield_obs_str
        return IDEObs(obs=obs_str).model_dump(), 100.0, done, success, {}

    def run(self, *, cmd: str, args: list[str] | None = None) -> int:
        """
        运行命令

        Args:
            cmd (str): 命令 | Command
            args (str): 参数 | Arguments

        Returns:
            int: 进程ID
        """
        self._assert_not_closed()
        # 使用subprocess.Popen运行指定命令
        # Run the specified command using subprocess.Popen
        proc = subprocess.Popen(
            [cmd] + args if args is not None else cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=self.current_dir,
            text=True,
        )
        self.procs[proc.pid] = proc

        def format_args(f_args: Any) -> str:
            if isinstance(f_args, list):
                return " ".join(arg.decode("utf-8") if isinstance(arg, bytes) else arg for arg in f_args)
            elif isinstance(f_args, bytes):
                return f_args.decode("utf-8")
            elif isinstance(f_args, str):
                return f_args
            else:
                raise ValueError(f"Unknow type of {f_args}")

        self._procs_working_description[proc.pid] = f"${self.current_dir}: {format_args(proc.args)}\n"
        return proc.pid

    def get_proc_res(self, *, pid: int) -> tuple[str, bool, bool]:
        """
        通过pid获取对应进程的返回

        Args:
            pid (int): 进程ID | Process ID

        Returns:
            str : 进程返回 | Process return
            bool: 是否完成 | Whether it is completed
            bool: 是否成功 | Whether it is successful
        """
        if pid not in self._procs_result:
            # 如果未捕获过这个进程的结果，则在此执行捕获逻辑
            self._procs_result[pid] = self.capture_proc_stdout(proc=self.procs[pid])

        # 检查进程是否完成 / Check if process is done
        proc = self.procs[pid]
        is_done = proc.poll() is not None

        # 如果进程已完成且还没有记录最终状态，保存它 / If process is done and final status not recorded, save it
        if is_done and pid not in self._procs_final_success:
            self._procs_final_success[pid] = proc.returncode == 0

        # 获取默认的 success 值：如果进程已完成，使用保存的最终状态 / Get default success value
        default_success = self._procs_final_success.get(pid, True)

        # 从迭代器获取结果，使用正确的默认值 / Get result from iterator with correct default
        res, success = next(self._procs_result[pid], ("Command Finished\n", default_success))
        self._procs_working_description[pid] += res

        return res, is_done, success

    @staticmethod
    def capture_proc_stdout(proc: subprocess.Popen, timeout: float = 0.1) -> Iterator[tuple[str, bool]]:
        """
        使用select模块对进程输出进行监听，如果有输出，则返回。如果是持续输出，比如tail -f 则通过yield返回。

        Args:
            proc (subprocess.Popen): 被监听输出的进程
            timeout (float): 监听周期时间设置，避免技术消耗CPU，同时不释放线程

        Returns:
            Iterator[tuple[str, bool]]: 返回结果。捕获的文本与当前执行是否有异常
        """
        if not proc.stdout:
            yield "进程未指定stdout，无法正常捕获", False  # pragma: no cover
        if proc.poll() is not None:
            # 进程已经完成 / Process has finished
            # 使用退出码判断成功与否，而不仅仅依赖 stderr / Use exit code to determine success, not just stderr
            returncode = proc.returncode
            success = returncode == 0

            stdout = proc.stdout.read() if proc.stdout else None
            stderr = proc.stderr.read() if proc.stderr else None

            # 合并 stdout 和 stderr，优先显示 stderr / Merge stdout and stderr, prioritize stderr
            if stderr and stdout:
                # 如果同时有 stderr 和 stdout，合并输出 / If both stderr and stdout exist, merge output
                yield stderr + stdout, success
            elif stderr:
                yield stderr, success
            elif stdout:
                yield stdout, success
            else:
                yield "no output", success
        else:
            # 进程正在执行中
            while True:
                if proc.poll() is None:
                    rlist, _, _ = select.select([proc.stdout], [], [], timeout)
                    err_rlist, _, _ = select.select([proc.stderr], [], [], timeout)
                    if err_rlist:
                        stderr = proc.stderr.read() if proc.stderr else None
                        if stderr:
                            # 进程还在运行时，stderr 通常表示错误 / stderr usually indicates error when process is running
                            yield stderr, False
                    if rlist:
                        stdout = proc.stdout.read() if proc.stdout else None
                        if stdout:
                            # 进程还在运行时，stdout 通常表示正常输出 / stdout usually indicates normal output when process is running
                            yield stdout, True
                    else:
                        continue  # pragma: no cover
                else:
                    # 程序正常终止
                    break

    def reset(
        self,
        *,
        seed: int | None = None,
        options: dict[str, Any] | None = None,
    ) -> tuple[IDEObs, dict[str, Any]]:
        """
        重置环境

        目前仅实现将工作目录切换为初始化目录

        Args:
            seed:
            options:

        Returns:
            IDEObs: 观察结果 | Observation result
            dict[str, Any]: 额外信息 | Additional information
        """
        self._assert_not_closed()
        self.change_dir(path=self.work_dir)
        return IDEObs(obs="Reset environment"), {}

    def render(self) -> str:  # type: ignore
        """
        以文本的形式来返回当前所有process的命令与结果

        当前默认仅渲染倒数后三个命令与其结果

        Returns:
            str: 渲染结果 | Render result
        """
        self._assert_not_closed()
        render_frames = []
        proc_keys = list(self.procs.keys())  # Convert to list
        for pid in proc_keys[-3:]:
            render_frames.append(self._procs_working_description[pid])
        return "\n".join(render_frames) if render_frames else f"${self.current_dir}: "

    def close(self) -> None:
        """
        关闭环境

        Args:
            self: The current instance of the class.

        Returns:
            None
        """
        self._is_closing = True
        for proc in self.procs.values():
            if proc is not None:
                if proc.stdin:
                    proc.stdin.close()
                try:
                    proc.send_signal(signal.SIGINT)
                    proc.wait(timeout=10)
                except subprocess.TimeoutExpired:  # pragma: no cover
                    proc.kill()  # pragma: no cover
        self._procs_result.clear()
        self._procs_final_success.clear()
        self._procs_working_description.clear()
        self._is_closed = True
        self._is_closing = False

    def _assert_not_closed(self) -> bool:
        """
        Assert that the environment is not closed.

        Returns:
            bool: True if the environment is not closed, False otherwise.
        """
        if self._is_closed:
            raise ValueError("Environment is closed.")  # pragma: no cover
        return True

    def change_dir(self, *, path: str) -> None:
        """
        更改当前目录

        Args:
            path (str): 新的目录路径 | New directory path

        Returns:
            None
        """
        # 判断path是否是一个存在的合法目录。同时判断path是否是self.work_dir的子目录
        # Determine if path is an existing valid directory. Also determine if path is a subdirectory of self.work_dir
        # 如果是，则更改当前目录为path
        # If so, change the current directory to path
        # 使用 realpath 来解析路径，包括 ~ 和符号链接
        path = os.path.expanduser(path)
        real_path = os.path.realpath(path)
        real_work_dir = os.path.realpath(self.work_dir)

        # 确保 real_path 是一个存在的目录
        if os.path.exists(real_path) and os.path.isdir(real_path):
            # 使用 commonpath 来判断是否是子目录
            try:
                common_path = os.path.commonpath([real_path, real_work_dir])
                if common_path == real_work_dir:
                    self.current_dir = real_path
                else:
                    raise ValueError(f"The path {path} is not a subdirectory of the working directory {self.work_dir}.")
            except ValueError as e:
                # commonpath 抛出 ValueError 如果路径列表为空或路径不在同一个驱动器上
                raise ValueError(
                    f"The path {path} is not a subdirectory of the working directory {self.work_dir}.",
                ) from e
        else:
            raise ValueError(f"Invalid path: {path}")
