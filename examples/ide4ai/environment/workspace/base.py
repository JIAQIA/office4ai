# filename: base.py
# @Time    : 2024/4/18 10:48
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import json
import os.path
import select
import signal
import subprocess
import threading
import time
from abc import ABC, abstractmethod
from collections import OrderedDict
from collections.abc import Callable, Sequence
from io import BufferedReader
from json import JSONDecodeError
from pathlib import Path
from typing import Any, ClassVar, Literal, cast

import gymnasium as gym
from cachetools import TTLCache
from gymnasium.core import RenderFrame
from loguru import logger
from pydantic import AnyUrl
from typing_extensions import SupportsFloat

from ide4ai.dtos.base_protocol import LSPResponseMessage
from ide4ai.dtos.diagnostics import DocumentDiagnosticReport, PreviousResultId, WorkspaceDiagnosticReport
from ide4ai.dtos.workspace_edit import LSPWorkspaceEdit
from ide4ai.environment.workspace.model import TextModel
from ide4ai.environment.workspace.schema import Position, Range, SearchResult, SingleEditOperation, TextEdit
from ide4ai.schema import ACTION_CATEGORY_MAP, IDEAction, IDEObs
from ide4ai.utils import is_subdirectory, list_directory_tree, render_symbols


class BaseWorkspace(gym.Env, ABC):
    """
    编辑工程文件的工作区

    1. 逐步支持LSP（当前尚未完全支持）
    2. 当前每个Workspace仅支持单一root_dir与project_name

    Attributes:
        name (str): The name of the environment.
        metadata (dict[str, Any]): The metadata of the environment.
        root_dir (str): The root directory of the workspace.
        project_name (str): The name of the project.
        models (list[Models]): The models of workspace. Models are at the heart of Monaco editor. It's what you interact
            with when managing content. A model represents a file that has been opened. This could represent a file that
            exists on a file system, but it doesn't have to. For example, the model holds the text content, determines
            the language of the content, and tracks the edit history of the content.
    """

    name: ClassVar[str]
    metadata: dict[str, Any] = {"render_modes": ["ansi"]}

    def __init__(
        self,
        root_dir: str,
        project_name: str,
        render_with_symbols: bool = True,
        max_active_models: int = 3,
        enable_simple_view_mode: bool = False,
        header_generators: dict[str, Callable[["BaseWorkspace", str], str]] | None = None,
        shortcut_commands: dict[str, list[str]] | None = None,
        diagnostics_timeout: float = 10.0,
        *args: Any,
        **kwargs: Any,
    ) -> None:
        super().__init__(*args, **kwargs)
        if not os.path.exists(root_dir):
            raise ValueError(f"项目根目录 {root_dir} 不存在")
        if not os.path.realpath(root_dir):
            raise ValueError("必须使用绝对路径作用项目根目录参数")
        self.root_dir = root_dir
        self.expand_folders: set[str] | Literal["all"] = set()
        self.project_name = project_name
        self.models: list[TextModel] = []
        self._active_models: OrderedDict[str, TextModel] = OrderedDict()
        self.lsp_stdout_mutex = threading.Lock()
        self.lsp_stdin_mutex = threading.Lock()
        self.lsp_mutex = threading.Lock()
        self.lsp: subprocess.Popen[bytes] | None = None
        self._lsp_msg_id = 1
        self._max_active_models = max_active_models
        self._render_with_symbols = render_with_symbols
        self._enable_simple_view_mode = enable_simple_view_mode
        self.lsp_output_monitor_thread: threading.Thread | None = None
        # LSP输出缓冲区，用于累积未完整的消息 / LSP output buffer for accumulating incomplete messages
        self._lsp_buffer: str = ""
        # 诊断信息拉取超时时间（秒）/ Diagnostics pull timeout in seconds
        self._diagnostics_timeout = diagnostics_timeout
        # 请注意，对以下两个缓存的操作，需要在with self.lsp_mutex 上下文中进行，保证线程安全
        # 其中key一般使用lsp Notification的method字段，因为对于每个method，我们只需要处理最后一次的通知。但有时候也会使用method+uri的方式，
        # 比如diagnostic。
        self.lsp_server_notifications: TTLCache = TTLCache(maxsize=1000, ttl=300)
        # 对于发起的request，我们需要等待response，因此需要缓存response，key值是request_id
        self.lsp_server_response: TTLCache = TTLCache(maxsize=1000, ttl=300)
        # 初始化动作空间与观察空间
        self.action_space = gym.spaces.Dict(
            {
                "category": gym.spaces.Discrete(2),
                "action_name": gym.spaces.Text(100),
                "action_args": gym.spaces.Text(1000),
            },
        )
        self._action_category_map = ACTION_CATEGORY_MAP
        self.observation_space = gym.spaces.Dict(
            {
                "created_at": gym.spaces.Text(100),
                "obs": gym.spaces.Text(100000),
            },
        )
        self.launch_lsp()
        self._initial_lsp()
        self._is_closing = False
        self._is_closed = False
        self.header_generators: dict[str, Callable[[BaseWorkspace, str], str]] | None = header_generators
        self.shortcut_commands: dict[str, list[str]] | None = shortcut_commands

    def get_lsp_msg_id(self) -> int:
        """
        Get the next available message id for the Language Server Protocol (LSP) server.

        Returns:
            int: The next available message id.
        """
        if not self._lsp_msg_id:
            self._lsp_msg_id = 1
        self._lsp_msg_id += 1
        return self._lsp_msg_id

    def launch_lsp(self) -> None:
        """
        Launch the Language Server Protocol (LSP) server. Relaunch the LSP server if it is already running.

        Returns:
            None
        """
        with self.lsp_mutex:
            if self.lsp:
                self.kill_lsp()
            self.lsp = self._launch_lsp()
            self._start_lsp_monitor_thread()

    def send_lsp_msg(
        self,
        method: str,
        params: dict[str, Any] | None = None,
        message_id: int | None = None,
    ) -> str | None:
        """
        Send a message to the Language Server Protocol (LSP) server.

        通过stdin发送消息给LSP server

        Raises:
            ValueError: If the LSP server is not running.

        Args:
            method (str): The method of the message.
            params (dict[str, Any]): The parameters of the message. This parameter is optional and defaults to None.
            message_id (int): The message id. This parameter is optional and defaults to None. Request messages should
                include a message id. Notification messages should not include a message id.

        Returns:
            Optional[str]: The response of the LSP server.
        """
        if not self.lsp:
            raise ValueError("LSP server is not running.")
        msg: dict = {
            "jsonrpc": "2.0",
            "method": method,
            "params": params or {},
        }
        if message_id is not None:
            msg["id"] = message_id
        # 转换为JSON字符串
        msg_str = json.dumps(msg)
        # 计算 content_length
        content_length = len(msg_str.encode("utf-8"))
        # 发送请求
        with self.lsp_stdin_mutex:
            full_message = f"Content-Length: {content_length}\r\n\r\n{msg_str}"
            if self.lsp.stdin:
                self.lsp.stdin.write(
                    full_message.encode("utf-8"),
                )  # LSP进程以bytes模式打开，因为LSP协议也是按照bytes进行传输的与长度计算
                self.lsp.stdin.flush()
        return self.read_response(message_id) if message_id else None

    def _start_lsp_monitor_thread(self) -> None:
        """
        Start the thread to monitor the output of the Language Server Protocol (LSP) server.

        Returns:
            None
        """
        if self.lsp_output_monitor_thread and self.lsp_output_monitor_thread.is_alive():
            return
        self.lsp_output_monitor_thread = threading.Thread(target=self.__read_lsp_output, daemon=True)
        self.lsp_output_monitor_thread.start()

    @abstractmethod
    def _initial_lsp(self) -> None:
        """
        初始化 LSP 服务

        Returns:

        """
        ...

    def __read_lsp_output(self) -> None:
        """
        Read the output of the Language Server Protocol (LSP) server.

        注意要设置合理的退出机制，在self.lsp停止的情况下，退出循环
        同时修改数据时需要使用self.lsp_mutex上锁

        Returns:
            None
        """
        while True:
            if self.lsp and self.lsp.poll() is None and self.lsp.stdout:
                with self.lsp_stdout_mutex:
                    # 检查LSP进程状态 / Check LSP process status
                    if not self.lsp or not self.lsp.stdout or self.lsp.poll() is not None:
                        break
                    # 再次调用select，确保在获取锁的过程中不出现其它意外 / Call select again to ensure no issues during lock acquisition
                    # 检查是否有数据可读 / Check if there is data to read
                    rlist, _, _ = select.select([self.lsp.stdout], [], [], 0.1)
                    if not rlist:
                        continue

                    logger.info("获取到LSP服务器返回数据 / Got LSP server response data")

                    # 持续读取所有可用数据，直到没有完整消息为止 / Continuously read all available data until no complete messages
                    while True:
                        # 尝试从缓冲区解析完整消息 / Try to parse complete messages from buffer
                        message_parsed = self._try_parse_one_message()
                        if not message_parsed:
                            # 缓冲区中没有完整消息，尝试读取更多数据 / No complete message in buffer, try to read more data
                            # 使用非阻塞select检查是否还有数据 / Use non-blocking select to check if more data is available
                            if not self.lsp.stdout:
                                break
                            rlist, _, _ = select.select([self.lsp.stdout], [], [], 0)
                            if not rlist:
                                # 没有更多数据可读 / No more data to read
                                break

                            # 使用read1进行单次读取，不会等待填满缓冲区 / Use read1 for single read, won't wait to fill buffer
                            try:
                                # read1(n) 最多读取n字节，但不会阻塞等待填满n字节 / read1(n) reads at most n bytes without blocking to fill
                                # subprocess.Popen[bytes].stdout 实际是 BufferedReader 类型 / Actually BufferedReader type
                                stdout = cast(BufferedReader, self.lsp.stdout)
                                chunk = stdout.read1(4096)
                                if not chunk:
                                    break
                                self._lsp_buffer += chunk.decode("utf-8")
                            except Exception as e:
                                logger.error(f"读取LSP输出时出错 / Error reading LSP output: {e}")
                                break
            else:
                # lsp已经停止 / LSP has stopped
                break

    def _try_parse_one_message(self) -> bool:
        """
        尝试从缓冲区解析一条完整的LSP消息 / Try to parse one complete LSP message from buffer

        Returns:
            bool: 如果成功解析了一条消息返回True，否则返回False / True if a message was parsed, False otherwise
        """
        # LSP消息格式：Content-Length: ...\r\n\r\n{json}
        if not self._lsp_buffer:
            return False

        # 查找消息头 / Find message header
        if not self._lsp_buffer.startswith("Content-Length:"):
            # 查找下一个消息头 / Find next message header
            header_start = self._lsp_buffer.find("Content-Length:")
            if header_start == -1:
                # 没有完整的消息头，清空无效数据 / No complete header, clear invalid data
                self._lsp_buffer = ""
                return False
            # 丢弃消息头之前的数据 / Discard data before header
            self._lsp_buffer = self._lsp_buffer[header_start:]

        # 解析内容长度 / Parse content length
        header_end = self._lsp_buffer.find("\r\n\r\n")
        if header_end == -1:
            # 消息头不完整 / Incomplete header
            return False

        header = self._lsp_buffer[:header_end]
        length_line = header.split("\r\n")[0]
        try:
            length = int(length_line.split(":")[1].strip())
        except (ValueError, IndexError) as e:
            logger.error(f"解析Content-Length失败 / Failed to parse Content-Length: {e}")
            # 跳过这个无效的消息头 / Skip this invalid header
            self._lsp_buffer = self._lsp_buffer[header_end + 4 :]
            return False

        # 检查是否有完整的消息体 / Check if complete message body is available
        message_start = header_end + 4  # 跳过"\r\n\r\n" / Skip "\r\n\r\n"
        if len(self._lsp_buffer) < message_start + length:
            # 消息体不完整 / Incomplete message body
            return False

        # 提取完整消息 / Extract complete message
        message_body = self._lsp_buffer[message_start : message_start + length]
        self._lsp_buffer = self._lsp_buffer[message_start + length :]  # 剩余数据 / Remaining data

        logger.info(f"解析到一条完整LSP消息 / Parsed one complete LSP message: {message_body[:100]}...")

        # 处理消息 / Process message
        try:
            response_data = json.loads(message_body)
            with self.lsp_mutex:
                if "id" in response_data:
                    self.lsp_server_response[response_data["id"]] = message_body
                elif "method" in response_data:
                    params = response_data.get("params", {})
                    uri = str(params.get("uri")) if isinstance(params, dict) else "NotExists"
                    key = self.__construct_notification_key(response_data["method"], uri)
                    self.lsp_server_notifications[key] = message_body
        except JSONDecodeError as e:
            logger.error(f"JSON解析失败 / Failed to decode JSON: {e}, message: {message_body}")

        return True

    @staticmethod
    def __construct_notification_key(method: str, uri: str) -> str:
        """
        Construct a key for the notification cache.

        Args:
            method (str): The method of the notification.
            uri (str): The URI of the notification.

        Returns:
            str: The constructed key.
        """
        return f"{method}:{uri}" if method == "textDocument/publishDiagnostics" else method

    def read_response(self, request_id: int, timeout: float = 1) -> str | None:
        """
        Read the response of the Language Server Protocol (LSP) server.

        Args:
            request_id (int): The request id.
            timeout (int): The timeout value in seconds. This parameter is optional and defaults to 1.

        Returns:
            Optional[str]: The response of the LSP server.
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            if request_id in self.lsp_server_response:
                return cast(str, self.lsp_server_response.pop(request_id))
            time.sleep(0.1)
        return None

    def read_notification(self, method: str, uri: str, timeout: float = 0.05) -> str | None:
        """
        Read the notification of the Language Server Protocol (LSP) server.

        Args:
            method (str): The method of the notification.
            uri (str): The URI of the notification about resource.
            timeout (float): The timeout value in seconds. This parameter is optional and defaults to 1.

        Returns:
            Optional[str]: The notification of the LSP server.
        """
        start_time = time.time()
        notification_key = self.__construct_notification_key(method, uri)
        while time.time() - start_time < timeout:
            if notification_key in self.lsp_server_notifications:
                return cast(str, self.lsp_server_notifications.pop(notification_key))
            time.sleep(0.1)
        return None

    def pull_diagnostics(
        self,
        uri: str | None = None,
        previous_result_id: str | None = None,
        previous_result_ids: list[dict[str, str]] | None = None,
        timeout: float = 1.0,
    ) -> DocumentDiagnosticReport | WorkspaceDiagnosticReport | None:
        """
        主动拉取诊断信息 / Pull diagnostics actively

        支持两种模式：
        1. 文档诊断（Document Diagnostics）：当提供 uri 参数时，拉取单个文档的诊断信息
        2. 工作区诊断（Workspace Diagnostics）：当 uri 为 None 时，拉取整个工作区的诊断信息

        Supports two modes:
        1. Document Diagnostics: Pull diagnostics for a single document when uri is provided
        2. Workspace Diagnostics: Pull diagnostics for entire workspace when uri is None

        Args:
            uri (str | None): 文档的URI，如果为None则拉取工作区诊断 / Document URI, pull workspace diagnostics if None
            previous_result_id (str | None): 上一次文档诊断的结果ID / Previous result ID for document diagnostics
            previous_result_ids (list[dict[str, str]] | None): 上一次工作区诊断的结果ID列表 /
                Previous result IDs for workspace diagnostics
            timeout (float): 超时时间（秒）/ Timeout in seconds

        Returns:
            DocumentDiagnosticReport | WorkspaceDiagnosticReport | None: 诊断报告，如果失败则返回None /
                Diagnostic report, or None if failed

        Examples:
            # 拉取单个文档的诊断 / Pull diagnostics for a single document
            doc_diagnostics = workspace.pull_diagnostics(uri="file:///path/to/file.py")

            # 拉取整个工作区的诊断 / Pull diagnostics for entire workspace
            workspace_diagnostics = workspace.pull_diagnostics()

            # 使用上一次的结果ID进行增量拉取 / Incremental pull with previous result ID
            doc_diagnostics = workspace.pull_diagnostics(
                uri="file:///path/to/file.py",
                previous_result_id="previous-id-123"
            )
        """
        from ide4ai.dtos.diagnostics import (
            DocumentDiagnosticParams,
            RelatedFullDocumentDiagnosticReport,
            RelatedUnchangedDocumentDiagnosticReport,
            WorkspaceDiagnosticParams,
            WorkspaceDiagnosticReport,
        )

        msg_id = self.get_lsp_msg_id()

        # 发送请求但不等待响应 / Send request without waiting for response
        if not self.lsp:
            raise ValueError("LSP server is not running.")

        # 根据是否提供 uri 决定使用文档诊断还是工作区诊断 / Choose document or workspace diagnostics based on uri
        if uri is not None:
            # 文档诊断模式 / Document diagnostics mode
            params = DocumentDiagnosticParams(
                textDocument={"uri": uri},
                previousResultId=previous_result_id,
            ).model_dump(exclude_none=True)
            method = "textDocument/diagnostic"
        else:
            # 工作区诊断模式 / Workspace diagnostics mode
            params = WorkspaceDiagnosticParams(
                previousResultIds=[
                    PreviousResultId(uri=item["uri"], value=item["value"]) for item in (previous_result_ids or [])
                ],
            ).model_dump(exclude_none=True)
            method = "workspace/diagnostic"

        msg: dict = {
            "jsonrpc": "2.0",
            "method": method,
            "params": params,
            "id": msg_id,
        }
        msg_str = json.dumps(msg)
        content_length = len(msg_str.encode("utf-8"))

        with self.lsp_stdin_mutex:
            full_message = f"Content-Length: {content_length}\r\n\r\n{msg_str}"
            logger.info("准备发送拉取诊断信息的请求...")
            if self.lsp.stdin:
                self.lsp.stdin.write(full_message.encode("utf-8"))
                self.lsp.stdin.flush()

        # 使用 timeout 等待响应 / Wait for response with timeout
        start_time = time.time()
        while time.time() - start_time < timeout:
            logger.info("尝试获取诊断结果...")
            if msg_id in self.lsp_server_response:
                res = cast(str, self.lsp_server_response.pop(msg_id))
                try:
                    res_json = LSPResponseMessage.model_validate(json.loads(res))
                    if res_json.error:
                        logger.error(f"拉取诊断信息失败 / Failed to pull diagnostics: {res_json.error}")
                        return None

                    # 根据模式解析不同的响应类型 / Parse different response types based on mode
                    if uri is not None:
                        # 文档诊断响应 / Document diagnostics response
                        if isinstance(res_json.result, dict):
                            kind = res_json.result.get("kind")
                            if kind == "full":
                                return RelatedFullDocumentDiagnosticReport.model_validate(res_json.result)
                            elif kind == "unchanged":
                                return RelatedUnchangedDocumentDiagnosticReport.model_validate(res_json.result)
                        else:
                            logger.error(f"获取到非法诊断数据: {res_json.result}")
                            return None
                    else:
                        # 工作区诊断响应 / Workspace diagnostics response
                        if isinstance(res_json.result, dict):
                            return WorkspaceDiagnosticReport.model_validate(res_json.result)
                        else:
                            logger.error(f"获取到非法诊断数据: {res_json.result}")
                            return None

                except json.JSONDecodeError as e:
                    logger.error(f"解析诊断响应失败 / Failed to parse diagnostic response: {e}")
                    return None
            time.sleep(1)

        # 超时未收到响应 / Timeout without receiving response
        target = uri if uri else "workspace"
        logger.warning(f"拉取诊断信息超时 / Pull diagnostics timeout for {target}")
        return None

    @abstractmethod
    def _launch_lsp(self) -> subprocess.Popen[bytes]:
        """
        Launch the Language Server Protocol (LSP) server.

        Returns:
            subprocess.Popen[bytes]: The process of the LSP server.
        """
        ...

    def get_model(self, uri: str) -> TextModel | None:
        """
        Get a model by URI.

        Args:
            uri (str): The URI of the model to be retrieved.

        Returns:
            Optional[TextModel]: The model instance.
        """
        return next(filter(lambda m: m.uri == AnyUrl(uri), self.models), None)

    @property
    def active_models(self) -> list[TextModel]:
        """
        Get the active models.

        Returns:
            list[TextModel]: The active models.
        """
        return list(self._active_models.values())

    def active_model(self, model_id: str) -> None:
        """
        激活一个Model

        Args:
            model_id (str): Model的ID

        Returns:
            None
        """
        if len(self._active_models) >= self._max_active_models:
            self._active_models.popitem(last=False)  # Remove the oldest item
        if not any(m.m_id == model_id for m in self.models):
            raise ValueError(f"Model with ID {model_id} does not exist in models, open it first.")
        self._active_models[model_id] = next(filter(lambda m: m.m_id == model_id, self.models))
        self._active_models.move_to_end(model_id)  # Ensure the latest added/activated model is at the end

    def deactivate_model(self, model_id: str) -> None:
        """
        取消激活一个Model

        Args:
            model_id (str): Model的ID

        Returns:
            None
        """
        if model_id in self._active_models:
            del self._active_models[model_id]

    def clear_active_models(self) -> None:
        """
        清空所有激活的Model

        Returns:
            None
        """
        self._active_models.clear()

    def kill_lsp(self) -> None:
        """
        Kill the Language Server Protocol (LSP) server.

        Returns:
            None
        """
        # 关闭进程
        if self.lsp:
            if self.lsp.stdin:
                try:
                    self.lsp.stdin.close()
                except Exception as e:
                    logger.error(f"关闭LSP进程 stdin 时出错: {e}")

            # 尝试优雅关闭
            try:
                self.lsp.send_signal(signal.SIGINT)
                # 设置超时时间等待进程结束
                self.lsp.wait(timeout=2)  # 缩短到2秒，加快清理速度
                logger.info("LSP进程已优雅关闭 / LSP process gracefully terminated")
            except subprocess.TimeoutExpired:
                # 优雅关闭失败，强制终止
                logger.warning(
                    "LSP进程未响应SIGINT，尝试SIGTERM / LSP process didn't respond to SIGINT, trying SIGTERM"
                )
                try:
                    self.lsp.terminate()
                    self.lsp.wait(timeout=2)
                    logger.info("LSP进程已通过SIGTERM终止 / LSP process terminated via SIGTERM")
                except subprocess.TimeoutExpired:
                    # 强制终止也失败，使用SIGKILL
                    logger.warning(
                        "LSP进程未响应SIGTERM，使用SIGKILL强制终止 / LSP process didn't respond to SIGTERM, using SIGKILL"
                    )
                    self.lsp.kill()
                    self.lsp.wait(timeout=1)
                    logger.info("LSP进程已通过SIGKILL强制终止 / LSP process killed via SIGKILL")
            except Exception as e:
                # 捕获其他异常，确保进程被清理
                logger.error(f"关闭LSP进程时出错 / Error closing LSP process: {e}")
                try:
                    self.lsp.kill()
                    self.lsp.wait(timeout=1)
                except Exception:
                    pass  # 最后的尝试，忽略所有错误

            self.lsp = None

    def __del__(self) -> None:
        """

        Method Name: __del__

        Description:
        This method is called when the object is about to be destroyed and deallocated from memory. It invokes the
        `close()` method to perform any necessary cleanup operations.

        Parameters:
        self: The object instance on which the method is being called.

        Return Type:
        None

        """
        try:
            self.close()
        except Exception as e:
            # 在析构函数中捕获所有异常，避免影响垃圾回收
            # Catch all exceptions in destructor to avoid affecting garbage collection
            logger.error(f"析构时关闭环境出错 / Error closing environment in destructor: {e}")

    @abstractmethod
    def construct_action(self, action: dict) -> IDEAction:
        """
        Construct an instance of the IDEAction class from the provided action.

        Args:
            action (dict): A dictionary containing the action to be constructed.

        Returns:
            IDEAction: An instance of the IDEAction class representing the constructed action.
        """
        ...

    @abstractmethod
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
        # Format action to be compatible with the IDEAction class
        ...

    def reset(
        self,
        *,
        seed: int | None = None,
        options: dict[str, Any] | None = None,
    ) -> tuple[IDEObs, dict[str, Any]]:
        """
        重置环境
        将打开的文件关闭
        折叠所有展开的文件夹

        Args:
            seed: An integer value indicating the seed to be used for resetting. The seed is optional and defaults to
                None.

            options: A dictionary containing additional options for the reset operation. This parameter is optional and
                     defaults to None.

        Returns:
            A tuple containing two elements:
            1. An instance of the IDEObs class representing the initial
        """
        self._assert_not_closed()
        for m in self.models:
            m.dispose()
        self.models.clear()
        self.clear_active_models()
        return super().reset(seed=seed)

    @abstractmethod
    def render(self, *, verbose: bool = False) -> RenderFrame | list[RenderFrame] | None:
        """
        渲染当前工作区状态 | Render current workspace state

        Args:
            verbose (bool): 是否使用详细模式。True时返回包含Python包/模块描述的丰富信息，False时返回简化版本
                           | Whether to use verbose mode. True returns rich info with Python package/module descriptions,
                           False returns simplified version

        Returns:
            RenderFrame | list[RenderFrame] | None: 渲染结果 | Render result
        """
        ...

    def close(self) -> None:
        """
        关闭环境

        Args:
            self: The current instance of the class.

        Returns:
            None
        """
        # 防止重复关闭
        if self._is_closed or self._is_closing:
            return

        self._is_closing = True

        try:
            # 清理所有模型
            for m in self.models:
                try:
                    m.dispose()
                except Exception as e:
                    logger.error(f"清理模型时出错 / Error disposing model: {e}")
            self.models.clear()

            # 关闭LSP进程
            with self.lsp_mutex:
                self.kill_lsp()

            # 等待输出监控线程结束
            if self.lsp_output_monitor_thread and self.lsp_output_monitor_thread.is_alive():
                self.lsp_output_monitor_thread.join(timeout=3)  # 设置超时避免永久阻塞
                if self.lsp_output_monitor_thread.is_alive():
                    logger.warning(
                        "LSP输出监控线程未能在超时时间内结束 / LSP output monitor thread didn't finish in time"
                    )
        except Exception as e:
            logger.error(f"关闭环境时出错 / Error closing environment: {e}")
        finally:
            self._is_closed = True
            self._is_closing = False

    def _assert_not_closed(self) -> bool:
        """
        Assert that the environment is not closed.

        Returns:
            bool: True if the environment is not closed, False otherwise.
        """
        if self._is_closed:
            raise ValueError("Environment is closed.")
        return True

    @abstractmethod
    def open_file(self, *, uri: str) -> TextModel:
        """
        Open a file in the workspace.
        Initial a model instance, add it to self.models and active it

        Args:
            uri (str): The uri to the file to be opened.

        Returns:
            TextModel: The model instance representing the opened file.
        """
        ...

    def save_file(self, *, uri: str) -> None:
        """
        Save a file in the workspace.

        Args:
            uri (str): The URI of the file to be saved.

        Returns:
            None
        """
        self._assert_not_closed()
        tm = next(filter(lambda m: m.uri == AnyUrl(uri), self.models), None)
        if tm:
            # will_save reason
            # 1: Manually triggered, e.g. by the user pressing save, by starting debugging, or by an API call.
            # 2: Automatic after a delay
            # 3: When the editor lost focus
            self.send_lsp_msg("textDocument/willSave", {"textDocument": {"uri": uri}, "reason": 1})
            tm.save()
            self.send_lsp_msg(
                "textDocument/didSave",
                {"textDocument": {"uri": uri}, "text": tm.get_value()},
            )

    @abstractmethod
    def apply_edit(
        self,
        *,
        uri: str,
        edits: Sequence[SingleEditOperation | dict],
        compute_undo_edits: bool = False,
    ) -> tuple[list[TextEdit] | None, DocumentDiagnosticReport | None]:
        """
        Apply edits to a file in the workspace.

        Args:
            uri (str): The URI of the file to which the edits should be applied.
            edits (list[SingleEditOperation | dict]): The edits to be applied to the file.
            compute_undo_edits (bool): Whether to compute the undo edits. This parameter is optional and defaults to
                False.

        Returns:
            tuple[Optional[list[TextEdit]], Optional[DocumentDiagnosticReport]]:
                - The reverse edits that can be applied to undo the changes / 可用于撤销更改的反向编辑
                - Diagnostics result after editing / 编辑后的诊断结果
        """
        ...

    @abstractmethod
    def apply_workspace_edit(self, *, workspace_edit: LSPWorkspaceEdit) -> Any:
        """
        Apply a workspace edit to the workspace.

        Args:
            workspace_edit (LSPWorkspaceEdit): The workspace edit to be applied.

        Returns:
            Any: The result of applying the workspace edit.
        """
        ...

    @abstractmethod
    def rename_file(
        self,
        *,
        old_uri: str,
        new_uri: str,
        overwrite: bool | None = None,
        ignore_if_exists: bool | None = None,
    ) -> bool:
        """
        Rename a file in the workspace.

        Args:
            old_uri (str): 旧的URI信息
            new_uri (str): 新的URI信息
            overwrite (Optional[bool]): 如果文件存在是否覆盖。优先级高于ignore_if_exists
            ignore_if_exists (Optional[bool]): 如果文件存在是否忽略

        Returns:
            bool: 操作是否成功
        """
        ...

    @abstractmethod
    def delete_file(
        self,
        *,
        uri: str,
        recursive: bool | None = None,
        ignore_if_not_exists: bool | None = None,
    ) -> bool:
        """
        Delete a file in the workspace.

        Args:
            uri (str): The URI of the file to be deleted.
            recursive (Optional[bool]): Whether to delete the content recursively if a folder is denoted.
            ignore_if_not_exists (Optional[bool]): Whether to ignore the operation if the file does not exist.

        Returns:
            bool: True if the file was deleted successfully, False otherwise.
        """
        ...

    @abstractmethod
    def create_file(
        self,
        *,
        uri: str,
        overwrite: bool | None = None,
        ignore_if_exists: bool | None = None,
    ) -> tuple[TextModel | None, DocumentDiagnosticReport | None]:
        """
        Create a file in the workspace.

        Args:
            uri (str): The URI of the file to be created.
            overwrite (Optional[bool]): Whether to overwrite the target if it already exists.
            ignore_if_exists (Optional[bool]): Whether to ignore the operation if the file already exists.

        Returns:
            tuple[Optional[TextModel], Optional[DocumentDiagnosticReport]]:
                - The model instance representing the created file / 创建的文件模型实例
                - Diagnostics result after creation / 创建后的诊断结果
        """
        ...

    def close_file(self, *, uri: str) -> None:
        """
        Close a file in the workspace.

        Args:
            uri (str): The URI of the file to be closed.

        Returns:
            None
        """
        tm = next(filter(lambda m: m.uri == AnyUrl(uri), self.models), None)
        if tm:
            tm.dispose()
            self.deactivate_model(tm.m_id)
            self.models.remove(tm)
            self.send_lsp_msg("textDocument/didClose", {"textDocument": {"uri": uri}})

    def read_file(
        self,
        *,
        uri: str,
        with_line_num: bool = True,
        code_range: Range | None = None,
    ) -> str:
        """
        Read the content of a file in the workspace.

        Notes:
            if current workspace enable simple view mode, with_line_num will be ignored, the response of this function
            will always contain line number.

        Args:
            uri (str): The URI of the file to be read.
            with_line_num (bool): 是否带有行号。默认为True。
            code_range (Optional[Range]): The range of the code to be read. This parameter is optional and defaults to
                None.

        Returns:
            str: The content of the file.
        """
        tm: TextModel | None = next(filter(lambda m: m.uri == AnyUrl(uri), self.models), None)
        if tm:
            return (
                tm.get_view(with_line_num, code_range)
                if not self._enable_simple_view_mode
                else tm.get_simple_view(code_range)
            )
        else:
            tm = self.open_file(uri=uri)
            return (
                tm.get_view(with_line_num, code_range)
                if not self._enable_simple_view_mode
                else tm.get_simple_view(code_range)
            )

    def expand_folder(self, *, uri: str) -> str:
        """
        Expand a folder in the workspace.

        Args:
            uri (str): The URI of the folder to be expanded.

        Returns:
            str: The directory info after expanding the folder.
        """
        if not uri.startswith("file://"):
            raise ValueError("URI must start with 'file://'")
        folder_path = uri[7:]
        if not os.path.realpath(folder_path) or not os.path.exists(folder_path):
            raise ValueError(f"Invalid folder path: {folder_path}")
        if not is_subdirectory(folder_path, self.root_dir):
            raise ValueError(f"Folder path {folder_path} is not a subdirectory of the root directory {self.root_dir}")
        if self.expand_folders != "all":
            self.expand_folders.add(folder_path)
        return list_directory_tree(folder_path, include_dirs=self.expand_folders, recursive=True)

    def glob_files(
        self,
        *,
        pattern: str,
        path: str | None = None,
    ) -> list[dict]:
        """
        使用通配符模式匹配文件 / Match files using glob pattern

        支持通配符模式，如 "**/*.js" 或 "src/**/*.ts"
        按修改时间排序返回匹配的文件路径

        Supports wildcard patterns like "**/*.js" or "src/**/*.ts"
        Returns matched file paths sorted by modification time

        Args:
            pattern (str): 用于匹配文件的通配符模式 / Glob pattern for matching files
            path (str | None): 要搜索的目录。若未指定，将使用工作区根目录 /
                              Directory to search. If not specified, uses workspace root

        Returns:
            List[dict]: 匹配的文件列表，每个包含路径和修改时间 /
                       List of matched files with path and modification time

        Examples:
            # 查找所有 Python 文件 / Find all Python files
            workspace.glob_files(pattern="**/*.py")

            # 在特定目录查找 / Search in specific directory
            workspace.glob_files(pattern="*.js", path="src")

            # 递归查找 TypeScript 文件 / Recursively find TypeScript files
            workspace.glob_files(pattern="**/*.ts")
        """
        self._assert_not_closed()

        # 确定搜索路径 / Determine search path
        search_path = Path(path) if path else Path(self.root_dir)

        # 如果是相对路径，转换为相对于工作区根目录的绝对路径 / If relative path, convert to absolute path relative to workspace root
        if not search_path.is_absolute():
            search_path = Path(self.root_dir) / search_path

        # 验证路径是否存在 / Validate path exists
        if not search_path.exists():
            raise ValueError(f"搜索路径不存在 / Search path does not exist: {search_path}")

        # 确保搜索路径在工作区内 / Ensure search path is within workspace
        if not is_subdirectory(str(search_path), self.root_dir):
            raise ValueError(f"搜索路径必须在工作区根目录内 / Search path must be within workspace root: {search_path}")

        # 执行 glob 匹配 / Perform glob matching
        matched_files = []
        for file_path in search_path.glob(pattern):
            if file_path.is_file():
                try:
                    mtime = os.path.getmtime(file_path)
                    matched_files.append(
                        {
                            "uri": f"file://{file_path.absolute()}",
                            "path": str(file_path.relative_to(self.root_dir)),
                            "mtime": mtime,
                        },
                    )
                except (OSError, ValueError) as e:
                    # 跳过无法访问的文件 / Skip inaccessible files
                    logger.warning(f"无法访问文件 / Cannot access file {file_path}: {e}")
                    continue

        # 按修改时间降序排序（最新的在前）/ Sort by modification time descending (newest first)
        matched_files.sort(key=lambda x: cast(float, x["mtime"]), reverse=True)

        return matched_files

    def collapse_folder(self, *, uri: str) -> str:
        """
        Collapse a folder in the workspace.

        Args:
            uri (str): The URI of the folder to be collapsed.

        Returns:
            None

        Raises:
            ValueError: If the URI does not start with 'file://' or the folder path is not expanded.
        """
        if not uri.startswith("file://"):
            raise ValueError("URI must start with 'file://'")
        folder_path = uri[7:]
        if self.expand_folders == "all":
            self.expand_folders = set()
        if folder_path in self.expand_folders:
            self.expand_folders.remove(folder_path)
        else:
            raise ValueError(f"Folder path {folder_path} is not expanded")
        return list_directory_tree(folder_path, include_dirs=self.expand_folders, recursive=True)

    def get_file_symbols(self, *, uri: str, kinds: list[int]) -> str:
        """
        Get the symbols in a file in the workspace.

        Args:
            uri (str): The URI of the file to get the symbols from.
            kinds (list[int]): The kinds of symbols to get.

        Returns:
            str: The symbols in the file.
        """
        self._assert_not_closed()
        mid = self.get_lsp_msg_id()
        lsp_res = self.send_lsp_msg(
            "textDocument/documentSymbol",
            {"textDocument": {"uri": uri}},
            message_id=mid,
        )
        if lsp_res:
            res_model = LSPResponseMessage.model_validate(json.loads(lsp_res))
            if res_model.error:
                return res_model.error.message
            symbols = res_model.result
            res = render_symbols(cast(list[dict], symbols), kinds)
            return (
                res
                + "\n以上是文件的符号信息，每个信息后面跟着的是符号的位置信息，可以通过此位置信息与URI查询具体代码。"
            )
        else:
            return "获取文件符号失败"

    @abstractmethod
    def find_in_path(
        self,
        *,
        uri: str,
        query: str,
        search_scope: Range | list[Range] | None = None,
        is_regex: bool = False,
        match_case: bool = False,
        word_separator: str | None = None,
        capture_matches: bool = True,
        limit_result_count: int | None = None,
    ) -> list[SearchResult]:
        """
        在工作区中的文件或文件夹内查找查询字符串 / Find a query in a file or folder in the workspace.

        Args:
            uri (str): 要搜索的文件或文件夹的 URI。如果是文件夹，将递归搜索其中的所有文件 /
                      The URI of the file or folder to search in. If it's a folder, will recursively search all files within.
            query (str): 要搜索的查询字符串 / The query to search for.
            search_scope: 可选。指定搜索应在其中进行的范围或范围列表。仅当 uri 是文件时有效。如果未提供，
                则在整个文件范围内进行搜索 / Optional. The range or list of ranges where the search should be performed.
                Only valid when uri is a file. If not provided, the search will be performed in the full file range.
            is_regex: 可选。指定是否应将搜索字符串视为正则表达式。默认为 False /
                     Optional. Specifies whether the search string should be treated as a regular expression. Default is False.
            match_case: 可选。指定搜索是否应区分大小写。默认为 False /
                       Optional. Specifies whether the search should be case-sensitive. Default is False.
            word_separator: 可选。用于定义搜索中单词边界的分隔符。如果未提供，则所有字符都视为单词的一部分 /
                          Optional. The separator used to define word boundaries in the search. If not provided,
                          all characters are considered as part of a word.
            capture_matches: 可选。指定是否应在搜索结果中捕获匹配的文本内容。默认为 True /
                           Optional. Specifies whether the matched text should be captured in the search results. Default is True.
            limit_result_count: 可选。返回的搜索结果的最大数量。如果未提供，将返回所有匹配项 /
                              Optional. The maximum number of search results to return. If not provided, all matches will be returned.

        Returns:
            表示匹配结果的 SearchResult 对象列表。每个结果包含匹配的范围和文本（如果 capture_matches 为 True）/
            A list of SearchResult objects representing the matched results. Each result contains the matched range
            and text (if capture_matches is True).

        Raises:
            ValueError: 如果提供了无效的 URI 或搜索范围 / If an invalid URI or search scope is provided.

        Examples:
            # 在单个文件中搜索 / Search in a single file
            results = workspace.find_in_path(uri="file:///path/to/file.py", query="def")

            # 在文件夹中递归搜索 / Recursively search in a folder
            results = workspace.find_in_path(uri="file:///path/to/folder", query="TODO", match_case=True)

            # 使用正则表达式搜索 / Search with regex
            results = workspace.find_in_path(uri="file:///path/to/file.py", query=r"\\bclass\\s+\\w+", is_regex=True)
        """
        ...

    def replace_in_file(
        self,
        *,
        uri: str,
        query: str,
        replacement: str,
        search_scope: Range | list[Range] | None = None,
        is_regex: bool = False,
        match_case: bool = False,
        word_separator: str | None = None,
        compute_undo_edits: bool = False,
    ) -> tuple[list[TextEdit] | None, DocumentDiagnosticReport | None]:
        """
        在工作区的文件中替换查询字符串。
        Replace a query with a specified string in a file in the workspace.

        Args:
            uri (str): 要在其中执行替换的文件的 URI。| The URI of the file to perform the replacement in.
            query (str): 要搜索的查询字符串。| The query string to search for.
            replacement (str): 用于替换查询的字符串。| The string to replace the query with.
            search_scope: 可选。指定替换应在其中进行的范围或范围列表。如果未提供，则在整个模型范围内进行替换。|
                Optional. The range or list of ranges where the replacement should be performed. If not
                provided, the replacement will be performed in the full model range.
            is_regex: 可选。指定是否应将查询字符串视为正则表达式。默认为 False。|
                Optional. Specifies whether the query string should be treated as a regular expression. Default is False.
            match_case: 可选。指定替换是否应区分大小写。默认为 False。|
                Optional. Specifies whether the replacement should be case-sensitive. Default is False.
            word_separator: 可选。用于定义搜索和替换中单词边界的分隔符。如果未提供，则所有字符都视为单词的一部分。|
                Optional. The separator used to define word boundaries for the search and replacement. If
                not provided, all characters are considered as part of a word.
            compute_undo_edits: 可选。决定是否计算撤销编辑。默认为 False。|
                Optional. Specifies whether to compute the undo edits. Default is False.

        Returns:
            tuple[Optional[list[TextEdit]], Optional[DocumentDiagnosticReport]]:
                - 可用于撤销更改的反向编辑 / The reverse edits that can be applied to undo the changes
                - 编辑后的诊断结果 / Diagnostics result after editing
        """
        search_res = self.find_in_path(
            uri=uri,
            query=query,
            search_scope=search_scope,
            is_regex=is_regex,
            match_case=match_case,
            word_separator=word_separator,
        )
        if not search_res:
            return None, None
        edits = [SingleEditOperation(range=sr.range, text=replacement) for sr in search_res]
        undo_edits, diagnostics = self.apply_edit(uri=uri, edits=edits, compute_undo_edits=compute_undo_edits)
        return undo_edits, diagnostics

    def insert_cursor(self, *, uri: str, key: str, position: Position) -> str:
        """
        Inserts a cursor at the specified position in the given file.

        Args:
            uri (str): The URI of the file to insert the cursor into.
            key (str): The key associated with the cursor.
            position (Position): The position where the cursor should be inserted.

        Returns:
            str: The content near the inserted cursor.

        Raises:
            AssertionError: If the file is closed.
        """
        self._assert_not_closed()
        model = self.get_model(uri)
        if not model:
            model = self.open_file(uri=uri)
        near_content = model.insert_cursor(key=key, position=position)
        return near_content

    def delete_cursor(self, *, uri: str, key: str) -> str:
        """
        Args:
            uri (str): A string representing the URI of the file to perform the delete operation on.
            key (str): A string representing the key of the cursor to be deleted.

        Returns:
            str: A string representing the content near the deleted cursor position.

        Raises:
            AssertionError: If the database is closed.
            FileNotFoundError: If the specified file does not exist.

        """
        self._assert_not_closed()
        model = self.get_model(uri)
        if not model:
            model = self.open_file(uri=uri)
        near_content = model.delete_cursor(key=key)
        return near_content

    def clear_cursors(self, *, uri: str) -> str:
        """
        Clears all cursors in the given model.

        Args:
            uri (str): The URI of the model.

        Returns:
            str: The result of clearing the cursors.
        """
        self._assert_not_closed()
        model = self.get_model(uri)
        if not model:
            model = self.open_file(uri=uri)
        return model.clear_cursors()
