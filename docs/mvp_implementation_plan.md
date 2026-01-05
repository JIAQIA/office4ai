# Office4AI MVP 实现计划

> **目标**: 打通消息通信，实现至少一个 Word 工具端到端调用
> **时间范围**: 1-2 周
> **最后更新**: 2026-01-05

---

## 1. 当前状态总结

### 1.1 已完成 ✅

**office4ai (Python)**:
- ✅ Socket.IO Server 完整实现 (server.py, namespaces/word.py)
- ✅ Connection Manager 完整实现 (连接管理、文档映射)
- ✅ DTOs 完整定义 (word.py, excel.py, ppt.py, common.py)
- ✅ BaseMCPServer 架构完整
- ✅ 3 个 Word 事件处理器框架 (get:selectedContent, insert:text, replace:selection)

**office-editor4ai (TypeScript)**:
- ✅ Socket.IO Client 完整实现 (client.ts)
- ✅ Word 事件处理器完整实现 (word-handlers.ts, 3 个事件)
- ✅ 工具函数完整封装 (word-tools/)
- ✅ 单元测试覆盖 (client.test.ts, word-handlers.test.ts)
- ✅ 类型定义完整 (socketio-types.ts)

### 1.2 待完成 ❌

**office4ai (Python)**:
- ❌ BaseWorkspace 抽象基类
- ❌ OfficeWorkspace 具体实现
- ❌ Workspace 与 Socket.IO 的集成
- ❌ MCP Tool 实现（哪怕一个）
- ❌ 端到端测试脚本

**office-editor4ai (TypeScript)**:
- ❌ taskpane.ts 中集成 Socket.IO Client
- ❌ 与真实 Socket.IO 服务器的集成测试

---

## 2. MVP 目标

**核心目标**: 实现一个端到端的 Word 工具调用链路

```
测试脚本 → OfficeWorkspace → Socket.IO Server → Word Add-In → 返回结果
```

**最小可行工具**: `word_get_selected_content` (获取选中内容)

**成功标准**:
1. Python Workspace 能够启动 Socket.IO Server
2. Word Add-In 能够连接到 Server
3. 测试脚本能够调用 `word_get_selected_content`
4. 成功获取 Word 文档中的选中内容并返回

---

## 3. 实现阶段

### 阶段 1: Workspace 核心实现 (2-3 天)

#### 1.1 实现 BaseWorkspace 抽象基类

**文件**: `office4ai/environment/workspace/base.py`

```python
from abc import ABC, abstractmethod
from enum import Enum
from typing import Any
from pydantic import BaseModel

class OfficeAction(BaseModel):
    """统一动作格式"""
    category: Literal["word", "excel", "ppt"]
    action_name: str
    params: dict[str, Any]

class OfficeObs(BaseModel):
    """统一观察格式"""
    success: bool
    data: dict[str, Any]
    error: str | None = None
    metadata: dict[str, Any] = {}

class DocumentStatus(Enum):
    CONNECTED = "connected"
    DISCONNECTED = "disconnected"
    UNKNOWN = "unknown"

class BaseWorkspace(ABC):
    @abstractmethod
    async def execute(self, action: OfficeAction) -> OfficeObs:
        """执行统一动作接口"""
        pass

    @abstractmethod
    def get_document_status(self, document_uri: str) -> DocumentStatus:
        """获取文档状态"""
        pass

    @abstractmethod
    async def emit_to_document(
        self,
        document_uri: str,
        event: str,
        data: dict[str, Any]
    ) -> dict[str, Any]:
        """向指定文档发送 Socket.IO 事件"""
        pass
```

#### 1.2 实现 OfficeWorkspace

**文件**: `office4ai/environment/workspace/office_workspace.py`

```python
import asyncio
from typing import Any
from .base import BaseWorkspace, OfficeAction, OfficeObs, DocumentStatus
from .socketio.server import create_socketio_server
from .socketio.services.connection_manager import connection_manager

class OfficeWorkspace(BaseWorkspace):
    """Office Workspace 实现，集成 Socket.IO 服务器"""

    def __init__(self, host: str = "127.0.0.1", port: int = 3000):
        self.host = host
        self.port = port
        self.sio_server = None
        self._document_map: dict[str, str] = {}  # uri → socket_id
        self._running = False

    async def start(self):
        """启动 Socket.IO 服务器"""
        if self._running:
            return

        self.sio_server = create_socketio_server()
        # 创建 aiohttp app 并启动
        from aiohttp import web
        app = web.Application()
        self.sio_server.attach(app)

        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, self.host, self.port)
        await site.start()

        self._running = True
        print(f"✅ Workspace Socket.IO Server started on http://{self.host}:{self.port}")

    async def stop(self):
        """停止 Socket.IO 服务器"""
        if not self._running:
            return

        # 清理连接
        if self.sio_server:
            await self.sio_server.close()

        self._running = False
        print("✅ Workspace Socket.IO Server stopped")

    async def execute(self, action: OfficeAction) -> OfficeObs:
        """执行统一动作接口"""
        document_uri = action.params.get("document_uri")
        if not document_uri:
            return OfficeObs(
                success=False,
                data={},
                error="Missing document_uri in params"
            )

        # 检查文档状态
        status = self.get_document_status(document_uri)
        if status != DocumentStatus.CONNECTED:
            return OfficeObs(
                success=False,
                data={},
                error=f"Document not connected: {document_uri}"
            )

        # 构造事件名称
        event = f"{action.category}:{action.action_name}"

        # 发送 Socket.IO 事件
        try:
            result = await self.emit_to_document(document_uri, event, action.params)
            return OfficeObs(success=True, data=result)
        except Exception as e:
            return OfficeObs(success=False, data={}, error=str(e))

    def get_document_status(self, document_uri: str) -> DocumentStatus:
        """获取文档状态"""
        if connection_manager.is_document_active(document_uri):
            return DocumentStatus.CONNECTED
        return DocumentStatus.DISCONNECTED

    async def emit_to_document(
        self,
        document_uri: str,
        event: str,
        data: dict[str, Any]
    ) -> dict[str, Any]:
        """向指定文档发送 Socket.IO 事件"""
        socket_id = connection_manager.get_socket_by_document(document_uri)
        if not socket_id:
            raise ValueError(f"No socket found for document: {document_uri}")

        # 通过 Socket.IO 发送事件并等待响应
        # (具体实现需要配合 async Future 或回调机制)
        from .socketio.server import sio_server
        # 实现请求-响应模式
        ...
```

#### 1.3 集成测试脚本

**文件**: `tests/integration_tests/test_workspace_e2e.py`

```python
import asyncio
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.base import OfficeAction

async def test_workspace_startup():
    """测试 Workspace 启动"""
    workspace = OfficeWorkspace()
    await workspace.start()

    # 等待用户连接 Add-In
    input("请在 Word Add-In 中连接到 Workspace，然后按 Enter 继续...")

    # 检查连接状态
    status = workspace.get_document_status("file:///fake/document.docx")
    print(f"Document status: {status}")

    await workspace.stop()
    print("✅ Workspace startup test passed")

if __name__ == "__main__":
    asyncio.run(test_workspace_startup())
```

---

### 阶段 2: Socket.IO 请求-响应模式 (2-3 天)

#### 2.1 实现请求-响应机制

**问题**: Socket.IO 默认是异步的，需要实现请求-响应关联

**方案**: 使用 `asyncio.Future` 实现响应等待

**文件**: `office4ai/environment/workspace/socketio/server.py` (修改)

```python
from asyncio import Future
from typing import Dict

# 全局请求表
_pending_requests: Dict[str, Future] = {}

async def emit_with_response(sid: str, event: str, data: dict, timeout: float = 5.0):
    """发送 Socket.IO 事件并等待响应"""
    request_id = data.get("requestId")
    if not request_id:
        raise ValueError("Missing requestId in data")

    # 创建 Future 等待响应
    future = asyncio.Future()
    _pending_requests[request_id] = future

    # 发送事件
    sio.emit(event, data, to=sid)

    # 等待响应
    try:
        result = await asyncio.wait_for(future, timeout=timeout)
        return result
    except asyncio.TimeoutError:
        del _pending_requests[request_id]
        raise TimeoutError(f"Request {request_id} timed out")
    finally:
        _pending_requests.pop(request_id, None)

# 在 WordNamespace 中添加响应处理
async def on_word_response(self, sid: str, data: dict):
    """处理 Add-In 返回的响应"""
    request_id = data.get("requestId")
    if request_id in _pending_requests:
        future = _pending_requests[request_id]
        if not future.done():
            future.set_result(data)
```

#### 2.2 更新 WordNamespace 事件处理

**文件**: `office4ai/environment/workspace/socketio/namespaces/word.py` (修改)

```python
async def on_word_get_selectedContent(self, sid: str, data: dict) -> None:
    """处理获取选中内容请求"""
    from ..server import emit_with_response

    try:
        # 转发到 Add-In
        result = await emit_with_response(sid, "word:get:selectedContent", data)
        # 响应会通过 on_word_response 返回
    except TimeoutError:
        logger.error(f"word:get:selectedContent timed out")

async def on_word_response(self, sid: str, data: dict) -> None:
    """处理 word 响应（通用）"""
    # 通过全局 _pending_requests 处理
    ...
```

---

### 阶段 3: Word Add-In 集成 (1-2 天)

#### 3.1 在 taskpane.ts 中集成 Socket.IO Client

**文件**: `/word-editor4ai/src/taskpane/taskpane.ts` (修改)

```typescript
import { SocketIOClient } from '../socketio/client';
import { WordHandlers } from '../socketio/handlers/word-handlers';

let socketClient: SocketIOClient | null = null;

export async function initializeSocketIO() {
  if (socketClient) {
    return; // 已初始化
  }

  // 获取当前文档 URI
  const documentUri = await getDocumentUri();

  // 创建 Socket.IO 客户端
  const handlers = new WordHandlers();
  socketClient = new SocketIOClient(documentUri, handlers);

  // 连接到 Workspace
  await socketClient.connect();

  console.log('✅ Socket.IO Client connected to Workspace');
}

async function getDocumentUri(): Promise<string> {
  // 获取当前文档的 URI
  return await Word.run(async (context) => {
    const doc = context.document;
    await context.sync();
    return `file://${doc.url}`;
  });
}

// 在页面加载时初始化
$(document).ready(() => {
  initializeSocketIO().catch(console.error);
});
```

#### 3.2 更新 manifest.xml

**文件**: `/word-editor4ai/manifest.xml` (确保权限)

确保 Add-In 有足够的权限访问文档内容和网络。

---

### 阶段 4: 端到端测试 (1-2 天)

#### 4.1 完整端到端测试

**文件**: `tests/integration_tests/test_word_e2e.py`

```python
import asyncio
import pytest
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.base import OfficeAction

@pytest.mark.integration
async def test_word_get_selected_content():
    """端到端测试：获取 Word 选中内容"""
    # 1. 启动 Workspace
    workspace = OfficeWorkspace()
    await workspace.start()

    # 2. 等待 Add-In 连接
    await asyncio.sleep(2)  # 给 Add-In 时间连接

    # 3. 检查连接状态
    # (需要知道真实的 document_uri)
    document_uri = "file:///path/to/test.docx"

    status = workspace.get_document_status(document_uri)
    assert status == DocumentStatus.CONNECTED, "Add-In not connected"

    # 4. 执行动作
    action = OfficeAction(
        category="word",
        action_name="get:selectedContent",
        params={
            "requestId": "test-001",
            "documentUri": document_uri,
            "options": {
                "includeText": True,
                "includeImages": False,
                "includeTables": False
            }
        }
    )

    obs = await workspace.execute(action)

    # 5. 验证结果
    assert obs.success, f"Action failed: {obs.error}"
    assert "data" in obs.data
    print(f"✅ 获取选中内容成功: {obs.data}")

    # 6. 清理
    await workspace.stop()
```

#### 4.2 手动测试步骤

**文件**: `tests/integration_tests/MANUAL_TEST.md`

```markdown
# 手动测试步骤

## 1. 启动 Workspace Server

```bash
cd office4ai
python tests/integration_tests/manual/start_workspace.py
```

## 2. 在 Word 中加载 Add-In

1. 打开 Word (桌面版或网页版)
2. 加载 word-editor4ai Add-In
3. 等待 Socket.IO 连接成功（应该在 taskpane 中看到日志）

## 3. 在 Word 中选中一些文本

## 4. 运行测试脚本

```bash
python tests/integration_tests/manual/test_get_selected_content.py
```

## 5. 验证结果

- 测试脚本应该成功获取选中的文本内容
- Word Add-In 不应该崩溃
- Workspace Server 日志应该显示请求和响应
```

---

## 4. 关键文件清单

### 需要创建的文件

| 文件路径 | 用途 |
|---------|------|
| `office4ai/environment/workspace/base.py` | BaseWorkspace 抽象基类 |
| `office4ai/environment/workspace/office_workspace.py` | OfficeWorkspace 实现 |
| `tests/integration_tests/test_workspace_e2e.py` | Workspace 启动测试 |
| `tests/integration_tests/test_word_e2e.py` | Word 端到端测试 |
| `tests/integration_tests/MANUAL_TEST.md` | 手动测试步骤 |
| `tests/integration_tests/manual/start_workspace.py` | 手动启动脚本 |

### 需要修改的文件

| 文件路径 | 修改内容 |
|---------|---------|
| `office4ai/environment/workspace/socketio/server.py` | 添加请求-响应机制 |
| `office4ai/environment/workspace/socketio/namespaces/word.py` | 实现事件处理逻辑 |
| `word-editor4ai/src/taskpane/taskpane.ts` | 集成 Socket.IO Client |

---

## 5. 风险与缓解

| 风险 | 缓解措施 |
|------|---------|
| Socket.IO 连接失败 | 添加详细的错误日志和重试机制 |
| 请求-响应超时 | 实现合理的超时和重试逻辑 |
| Word API 调用失败 | 在 Add-In 端添加完善的错误处理 |
| CORS 问题 | 确保 Socket.IO Server 配置正确的 CORS 策略 |

---

## 6. 成功标准

- ✅ Workspace Server 能够成功启动
- ✅ Word Add-In 能够连接到 Server（日志显示连接成功）
- ✅ 测试脚本能够调用 `word:get:selectedContent`
- ✅ 成功获取并返回 Word 文档中的选中内容
- ✅ 所有测试通过（单元测试 + 集成测试）

---

## 7. 后续扩展 (不在 MVP 范围内)

- 实现 `word:insert:text` 和 `word:replace:selection`
- 实现 PPT 和 Excel 的对应功能
- 实现 MCP Tools 封装
- 添加更多错误处理和日志
- 实现文档状态验证（主动探测 + 被动监听）
