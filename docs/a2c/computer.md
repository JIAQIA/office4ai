# Computer 模块使用文档 / Computer Module Usage Documentation

A2C-SMCP Computer 模块提供了 **工具宿主侧（Computer 端）** 的实现，用于在本地或远程主机上统一管理多个 MCP Server，并通过 SMCP 协议将这些工具能力暴露给 Agent。

- **协议视角**：在 `a2c_rfc.md` 中，Computer 是房间内的「工具提供方」，与 Agent 通过 Server 中转交互。
- **实现视角**：`a2c_smcp.computer.Computer` 负责 MCP Server 生命周期、工具列表、桌面抽象（Desktop）、动态输入（inputs）等能力。

---

## 1. 概述 / Overview

Computer 模块主要包含以下能力：

- **MCP Server 管理**  
  通过 `MCPServerManager` 管理多个 MCP Server（stdio / streamable / sse），支持自动连接、重连和健康监测。

- **工具列表聚合与导出**  
  将 MCP Tool 转换为协议层的 `SMCPTool`，并通过 Socket.IO SMCP 客户端上报给 Server。

- **工具调用执行与二次确认**  
  通过 `Computer.aexecute_tool()` 统一发起 MCP 工具调用，支持 `auto_apply` 与自定义二次确认回调。

- **Desktop 抽象（window:// 资源）**  
  聚合 MCP Server 暴露的 window 资源，形成「桌面」视图，用于 IDE/Agent 的桌面浏览与截图流。

- **Inputs 管理与惰性解析**  
  使用 `MCPServerInput` 定义动态配置项，并通过 `InputResolver` 在需要时才解析，避免配置被提前「烧死」。

- **Socket.IO 绑定（与 SMCPComputerClient 协作）**  
  通过 weakref 引用 `SMCPComputerClient`，在工具列表或桌面资源变化时向 Server 上报相应事件。

---

## 2. 核心类型与结构 / Core Types and Structures

### 2.1 Computer 类 / `a2c_smcp.computer.Computer`

```python
from a2c_smcp.computer.computer import Computer
```

`Computer` 继承自泛型基类 `BaseComputer[PromptSession]`，面向 CLI 场景进行特化：

- 使用 `PromptSession` 作为交互 Session 类型；
- 集成 `InputResolver` / `ConfigRender`，对 MCP 配置中的占位符进行按需解析；
- 暴露异步方法用于：
  - 启动与关闭 MCP Server 管理器（`boot_up` / `shutdown` / `__aenter__` / `__aexit__`）；
  - 执行工具调用（`aexecute_tool`）；
  - 获取可用工具（`aget_available_tools`）；
  - 获取 Desktop 布局（`get_desktop`）；
  - 管理 inputs 及其值缓存。

### 2.2 BaseComputer 抽象基类 / `a2c_smcp.computer.base.BaseComputer`

```python
from a2c_smcp.computer.base import BaseComputer
```

`BaseComputer[S]` 抽象出与 Session 类型无关的能力约定，为不同交互形态（CLI / GUI / Web）提供统一接口：

- **MCP Server 生命周期管理**：
  - `boot_up()`
  - `aadd_or_aupdate_server()`
  - `aremove_server()`

- **Inputs 定义管理**：
  - `update_inputs()`
  - `add_or_update_input()`
  - `remove_input()`
  - `get_input()` / `list_inputs()`

- **Inputs 当前值（缓存）管理**：
  - `get_input_value()` / `set_input_value()`
  - `remove_input_value()` / `list_input_values()`
  - `clear_input_values()`

- **资源生命周期**：
  - `shutdown()`

具体实现（当前为 `Computer`）在子类中完成这些方法，并可以根据不同 UI/运行环境替换 Session 类型 `S`。

### 2.3 ToolCallRecord / 工具调用历史

```python
from a2c_smcp.computer.types import ToolCallRecord
```

`ToolCallRecord` 用于记录最近的工具调用历史（`Computer` 默认保留最近 10 条）：

```python
class ToolCallRecord(TypedDict):
    timestamp: str
    req_id: str
    server: str
    tool: str
    parameters: dict
    timeout: float | None
    success: bool
    error: str | None
```

在 `Computer.aexecute_tool()` 中，无论调用是否成功，都会记录一条 `ToolCallRecord` 以便后续审计与调试。

---

## 3. 快速开始 / Quick Start

### 3.1 最简初始化与启动 / Minimal Initialization and Boot

```python
import asyncio
from a2c_smcp.computer.computer import Computer
from a2c_smcp.computer.mcp_clients.model import MCPServerConfig, MCPServerStdioConfig, MCPServerStdioParameters

async def main() -> None:
    # 1. 定义一个最简单的 MCP Server 配置
    #    Define a minimal MCP server config
    server_cfg = MCPServerStdioConfig(
        name="example_server",
        disabled=False,
        forbidden_tools=[],
        tool_meta={},
        server_parameters=MCPServerStdioParameters(
            command="python",
            args=["-m", "my_mcp_server"],
            env=None,
            cwd=None,
            encoding="utf-8",
            encoding_error_handler="strict",
        ),
    )

    # 2. 创建 Computer 实例
    computer = Computer(
        name="my_computer",
        mcp_servers={server_cfg},
    )

    # 3. 启动 Computer（初始化 MCPServerManager 并按需连接 MCP 服务器）
    await computer.boot_up()

    # 4. 获取当前可用工具列表
    tools = await computer.aget_available_tools()
    print(f"Available tools: {len(tools)}")

    # 5. 关闭 Computer
    await computer.shutdown()

asyncio.run(main())
```

### 3.2 上下文管理 / Context Manager Usage

```python
from a2c_smcp.computer.computer import Computer

async def main() -> None:
    async with Computer(name="my_computer", mcp_servers={server_cfg}) as computer:
        tools = await computer.aget_available_tools()
        # 执行业务逻辑 / Do something with tools

asyncio.run(main())
```

---

## 4. MCP 配置与 Inputs / MCP Config and Inputs

### 4.1 MCPServerConfig 与 Inputs 的关系

- `MCPServerConfig`：描述单个 MCP Server 的连接方式与工具元数据（支持 stdio / streamable / sse）。
- `MCPServerInput`：描述 MCP 配置中的动态字段（如 API Key、工作目录等），通常包含：
  - `id`: 唯一标识；
  - `type`: `promptString` / `pickString` / `command`；
  - `default` 等字段。

`Computer` 在 `boot_up()` 时会：

1. 对每个 `MCPServerConfig` 执行 `model_dump(mode="json")`；
2. 通过 `ConfigRender` 和 `InputResolver` 将其中引用 `inputs` 的占位符按需解析；
3. 使用 `model_validate` 重建不可变的配置对象。

### 4.2 管理 Inputs 定义

```python
from a2c_smcp.computer.computer import Computer
from a2c_smcp.computer.mcp_clients.model import MCPServerPromptStringInput

# 创建 Computer
computer = Computer(name="my_computer")

# 更新 inputs 定义
prompt_input = MCPServerPromptStringInput(
    id="OPENAI_API_KEY",
    description="OpenAI API 密钥 / OpenAI API key",
    type="promptString",
    default=None,
    password=True,
)

computer.update_inputs({prompt_input})
```

### 4.3 管理 Inputs 当前值（缓存）

```python
# 设置某个 input 的值（写入缓存）
computer.set_input_value("OPENAI_API_KEY", "sk-xxx")

# 获取当前值
value = computer.get_input_value("OPENAI_API_KEY")

# 列出所有已解析值
values = computer.list_input_values()

# 删除指定值或清空缓存
computer.remove_input_value("OPENAI_API_KEY")
computer.clear_input_values()
```

> 详细行为（如惰性解析、default 使用规则）可参考 CLI 文档与 `InteractiveImpl` 实现。

---

## 5. 工具调用与二次确认 / Tool Calling and Second Confirmation

### 5.1 工具调用流程

`Computer.aexecute_tool()` 是对 MCP 工具调用的统一入口：

1. 通过 `MCPServerManager.avalidate_tool_call()` 确认工具存在并解析最终 `server_name` 与 `tool_name`；
2. 合并工具元数据 `ToolMeta`（具体配置优先，缺失字段回落 `default_tool_meta`）；
3. 判断合并后的 `auto_apply` 字段：
   - 若 `auto_apply is True`：直接调用 MCP 工具；
   - 否则：进入二次确认流程；
4. 无论结果成功或失败，都会写入一条 `ToolCallRecord` 到内部调用历史中。

### 5.2 自定义二次确认回调 / Custom Confirm Callback

```python
from a2c_smcp.computer.computer import Computer

# 定义二次确认回调 / Define confirm callback
def confirm_tool_call(req_id: str, server: str, tool: str, params: dict) -> bool:
    print(f"[Confirm] req_id={req_id}, server={server}, tool={tool}, params={params}")
    # 这里可以接入终端交互、GUI 弹窗或策略判断
    return True  # 返回 False 表示拒绝本次调用

computer = Computer(
    name="my_computer",
    mcp_servers={server_cfg},
    confirm_callback=confirm_tool_call,
)

result = await computer.aexecute_tool(
    req_id="example-req-id",
    tool_name="file_read",
    parameters={"path": "/tmp/demo.txt"},
    timeout=30,
)
```

当 `auto_apply` 未显式为 True 且提供了 `confirm_callback` 时，Computer 会在调用前触发此回调，并根据返回值决定是否真正执行工具。

---

## 6. Desktop 支持 / Desktop Support

### 6.1 window:// 资源与 Desktop

Computer 将 MCP Server 中的部分资源（通常以 `window://` URI 命名的资源）视为「窗口」，并将其组合为 Desktop：

- 通过 `a2c_smcp.utils.window_uri.is_window_uri()` 判断哪些资源参与 Desktop；
- 通过 `organize_desktop()` 将资源组织成桌面布局；
- 对应 SMCP 协议中的 `client:get_desktop` / `notify:desktop_refresh` 等事件。

### 6.2 获取 Desktop

```python
from a2c_smcp.computer.computer import Computer

# 获取当前 Desktop 信息
# Get current desktop information

desktops = await computer.get_desktop(size=10)
for d in desktops:
    print(d)
```

参数说明：

- `size`：限制返回的窗口数量；
- `window_uri`：若指定，则尝试只获取对应 WindowURI 的窗口信息。

> Desktop 的具体序列化形式与前端展示策略，可以根据项目需要进一步扩展。

---

## 7. 与 Socket.IO / SMCPComputerClient 的协作 / Collaboration with SMCPComputerClient

Computer 本身并不直接管理 Socket.IO 连接，而是通过 weakref 绑定 `SMCPComputerClient`：

- 在工具列表变化时，回调 `_on_manager_change()` 会调用 `client.emit_update_tool_list()`；
- 在 window 资源变化时，会触发 Desktop 刷新相关事件；
- 这样可以避免强引用循环，同时保持 Computer 对网络栈的解耦。

绑定流程示意：

```python
from a2c_smcp.computer.computer import Computer
from a2c_smcp.computer.socketio.client import SMCPComputerClient

computer = Computer(name="my_computer", mcp_servers={server_cfg})
client = SMCPComputerClient(computer=computer)

# 在 SMCPComputerClient 初始化过程中，会将自身赋值给 computer.socketio_client
# During SMCPComputerClient initialization, it will set computer.socketio_client = self
```

---

## 8. 最佳实践 / Best Practices

- **将 MCP 配置与 inputs 分离管理**  
  在配置文件中仅使用占位符，把实际值交给 `inputs` + `InputResolver` 管理，方便不同环境（开发/生产）切换。

- **合理使用 auto_apply 与 confirm_callback**  
  对高风险工具（文件写入、网络操作等）建议关闭 `auto_apply`，并实现交互式二次确认。

- **结合 Agent / Server 文档一起阅读**  
  - `docs/a2c_rfc.md`：协议级约定（事件名、房间模型、错误码等）；
  - `docs/server.md`：如何在后端挂载 SMCP Server Namespace；
  - `docs/agent.md`：如何在 Agent 侧消费 Computer 提供的工具能力。

- **利用调用历史进行审计与调试**  
  根据 `ToolCallRecord`，可以在 CLI 或日志系统中重放或分析最近的工具调用情况，辅助问题排查。

---

## 9. 后续扩展 / Future Work

- 支持更多 MCP 传输模式与认证方案；
- 提供 GUI/桌面环境下的 Computer 特化实现（替换 `Session` 类型）；
- 更丰富的 Desktop 布局与交互能力；
- 与 `a2c_smcp/computer/cli` 深度集成的使用示例与教程。
