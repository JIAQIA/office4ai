# filename: bash.py
# @Time    : 2025/10/29 12:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
"""
Bash 工具实现 | Bash Tool Implementation

提供在 IDE 环境中执行 Bash 命令的能力
Provides the ability to execute Bash commands in the IDE environment
"""

from typing import Any

from loguru import logger

from ide4ai.a2c_smcp.schemas import BashInput, BashOutput
from ide4ai.a2c_smcp.tools.base import BaseTool


class BashTool(BaseTool):
    """
    Bash 命令执行工具 | Bash Command Execution Tool

    通过 PythonIDE 的终端环境执行 Bash 命令
    Executes Bash commands through PythonIDE's terminal environment
    """

    # 输出字符数限制 | Output character limit
    MAX_OUTPUT_LENGTH = 30000

    @property
    def name(self) -> str:
        return "Bash"

    @property
    def description(self) -> str:
        return """在一个持久的 shell 会话中执行给定的 bash 命令，支持可选的超时设置，确保正确处理和安全措施。

重要：此工具用于终端操作，如 git、npm、docker 等。请勿将其用于文件操作（读取、写入、编辑、搜索、查找文件） - 请改用为此设计的专用工具。

在执行命令之前，请遵循以下步骤：
目录验证：
如果命令将创建新目录或文件，请首先使用 ls验证父目录是否存在且位置正确
例如，在运行 "mkdir foo/bar" 之前，首先使用 ls foo检查 "foo" 是否存在且是预期的父目录
命令执行：
始终用双引号引用包含空格的文件路径（例如，cd "path with spaces/file.txt"）
正确引用的示例：
cd "/Users/name/My Documents" (正确)
cd /Users/name/My Documents (错误 - 将失败)
python "/path/with spaces/script.py" (正确)
python /path/with spaces/script.py (错误 - 将失败)
确保正确引用后，执行命令。
捕获命令的输出。
使用说明：
command 参数是必需的。
您可以指定一个可选的超时时间（以毫秒为单位，最多 600000 毫秒 / 10 分钟）。如果未指定，命令将在 120000 毫秒（2 分钟）后超时。
如果您能用 5-10 个词清晰、简洁地描述此命令的作用，将会非常有帮助。
如果输出超过 30000 个字符，返回给您之前输出将被截断。
您可以使用 run_in_background参数在后台运行命令，这允许您在命令运行时继续工作。您可以使用 Bash 工具在输出可用时进行监控。使用此参数时，无需在命令末尾使用 '&'。
避免将 Bash 工具与 find、grep、cat、head、tail、sed、awk或 echo命令一起使用，除非明确指示或这些命令对于任务确实必要。相反，应始终优先使用这些命令的专用工具：
文件搜索：使用 Glob 工具（不要用 find 或 ls）
内容搜索：使用 Grep 工具（不要用 grep 或 rg）
读取文件：使用 Read 工具（不要用 cat/head/tail）
编辑文件：使用 Edit 工具（不要用 sed/awk）
写入文件：使用 Write 工具（不要用 echo >/cat <<EOF）
通信：直接输出文本（不要用 echo/printf）
当发出多个命令时：
如果命令相互独立且可以并行运行，请在单条消息中进行多次 Bash 工具调用。例如，如果需要运行 "git status" 和 "git diff"，请发送一条包含两次并行 Bash 工具调用的消息。
如果命令相互依赖且必须顺序运行，请使用单个 Bash 调用并用 '&&' 将它们链接在一起（例如，git add . && git commit -m "message" && git push）。例如，如果一个操作必须在另一个操作开始之前完成（如 mkdir 在 cp 之前，对于 git 操作 Write 在 Bash 之前，或 git add 在 git commit 之前），则顺序运行这些操作。
仅当需要顺序运行命令但不关心早期命令是否失败时，才使用 ';'
不要使用换行符分隔命令（在引用的字符串中使用换行符是可以的）
尝试通过使用绝对路径和避免使用 cd来在整个会话中保持当前工作目录。如果用户明确要求，可以使用 cd。
<good-example>
pytest /foo/bar/tests
</good-example>
<bad-example>
cd /foo/bar && pytest tests
</bad-example>"""

    @property
    def input_schema(self) -> dict[str, Any]:
        """返回 JSON Schema 格式的输入定义 | Return input definition in JSON Schema format"""
        return BashInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        执行 Bash 命令 | Execute Bash command

        Args:
            arguments: 包含命令和选项的参数字典 | Arguments dict containing command and options

        Returns:
            dict: 执行结果，包含输出、错误信息等 | Execution result with output, errors, etc.
        """
        # 验证输入参数 | Validate input arguments
        try:
            bash_input = self.validate_input(arguments, BashInput)
        except ValueError as e:
            err_info = f"参数验证失败 | Argument validation failed: {e}"
            logger.error(err_info)
            return BashOutput(
                success=False,
                output="",
                error=err_info,
                exit_code=1,
            ).model_dump()

        logger.info(f"执行 Bash 命令 | Executing Bash command: {bash_input.command}")

        try:
            # 构造 IDE action
            # action_args 必须是 str 或 list[str]，不能是 dict
            # action_args must be str or list[str], not dict
            action = {
                "category": "terminal",
                "action_name": bash_input.command,
                "action_args": bash_input.args
                if bash_input.args is not None
                else "",  # 使用传入的 args 或空字符串 | Use provided args or empty string
            }

            # 如果设置了超时，转换为秒 | Convert timeout to seconds if set
            if bash_input.timeout:
                # IDE 的 timeout 是秒，MCP 的是毫秒 | IDE timeout is in seconds, MCP is in milliseconds
                timeout_seconds = bash_input.timeout / 1000
                # 注意：这里可能需要根据实际的 IDE 实现来调整超时设置
                # Note: May need to adjust timeout setting based on actual IDE implementation
                logger.debug(f"设置超时 | Setting timeout: {timeout_seconds}s")

            # 执行命令 | Execute command
            obs, reward, done, success, info = self.ide.step(action)

            # 获取输出并进行截断处理 | Get output and truncate if necessary
            raw_output = str(obs.get("obs", ""))
            output_text = raw_output
            truncated = False

            if len(raw_output) > self.MAX_OUTPUT_LENGTH:
                output_text = raw_output[: self.MAX_OUTPUT_LENGTH]
                truncated = True
                logger.warning(
                    f"输出被截断 | Output truncated: {len(raw_output)} -> {self.MAX_OUTPUT_LENGTH} 字符 | characters",
                )

            # 构造输出 | Construct output
            output = BashOutput(
                success=bool(success),
                output=output_text,
                error=str(obs.get("error", "")) if "error" in obs else None,
                exit_code=info.get("exit_code"),
                metadata={
                    "reward": float(reward),
                    "done": done,
                    "description": bash_input.description,
                    "run_in_background": bash_input.run_in_background,
                    "truncated": truncated,
                    "original_length": len(raw_output) if truncated else None,
                },
            )

            logger.info(
                f"命令执行完成 | Command execution completed: success={output.success}, exit_code={output.exit_code}",
            )

            return output.model_dump()

        except Exception as e:
            logger.exception(f"执行命令时发生错误 | Error executing command: {e}")

            # 返回错误结果 | Return error result
            error_output = BashOutput(
                success=False,
                output="",
                error=str(e),
                exit_code=-1,
                metadata={"exception_type": type(e).__name__},
            )

            return error_output.model_dump()
