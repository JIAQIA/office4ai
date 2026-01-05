# Office4AI Workspace 实现计划

> **版本**: 2.0.0
> **创建日期**: 2026-01-05
> **状态**: 精简版

---

## 1. 项目定位

**Office4AI** 是一个专为 AI Agent 设计的 Office 文档管理环境，通过 **MCP 协议** 暴露工具能力，底层通过 **Workspace Socket.IO** 与 Office Add-In 通信。

### 1.1 核心架构

```
┌─────────────────────────────────────────────────────────────┐
│                     AI Agent / User                         │
└──────────────────────────┬──────────────────────────────────┘
                           │ MCP Protocol (stdio/sse/http)
                           ↓
┌────────────────────────────────────────────────────────────┐
│              Office4AI MCP Server (Python)                 │
│                                                            │
│  ┌────────────────────────────────────────────────────┐    │
│  │  MCP Tools Layer                                   │    │
│  │  - word_get_selected_content                       │    │
│  │  - word_insert_text                                │    │
│  │  - ppt_insert_text                                 │    │
│  │  - excel_set_cell_value                            │    │
│  └──────────────────┬─────────────────────────────────┘    │
│                     │                                      │
│  ┌──────────────────▼─────────────────────────────────┐    │
│  │  Workspace (统一动作接口)                           │    │
│  │  execute(category, action_name, params)            │    │
│  └──────────────────┬─────────────────────────────────┘    │
│                     │                                      │
│  ┌──────────────────▼─────────────────────────────────┐    │
│  │  Workspace Socket.IO Server                        │    │
│  │  - /word namespace                                 │    │
│  │  - /ppt namespace                                  │    │
│  │  - /excel namespace                                │    │
│  └───────────────────┬────────────────────────────────┘    │
└──────────────────────┼─────────────────────────────────────┘
                       │ Socket.IO (WebSocket)
                       ↓
┌──────────────────────────────────────────────────────────┐
│              Office Add-In Clients (Word/Excel/PPT)       │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐    │
│  │ Word Add-In  │  │ PPT Add-In   │  │ Excel Add-In │    │
│  │ Socket.IO    │  │ Socket.IO    │  │ Socket.IO    │    │
│  │ Client       │  │ Client       │  │ Client       │    │
│  └──────────────┘  └──────────────┘  └──────────────┘    │
└──────────────────────────────────────────────────────────┘
```

### 1.2 协议分离

| 协议 | 用途 | 方向 |
|------|------|------|
| **MCP** | AI Agent 调用 Office 工具 | Agent → Office4AI |
| **Workspace Socket.IO** | 与 Office Add-In 通信 | Office4AI ↔ Add-In |

> **注意**：Workspace Socket.IO 与 A2C Socket.IO 是完全独立的系统，不要混淆。

---

## 2. 核心设计

### 2.1 OfficeAction (统一动作接口)

```python
class OfficeAction(BaseModel):
    """统一动作格式，通过 Workspace.execute() 执行"""
    category: Literal["word", "excel", "ppt"]      # Office 应用类型
    action_name: str                                 # 操作名称 (如 "insert:text")
    params: dict[str, Any]                           # 操作参数
```

### 2.2 BaseWorkspace 接口

```python
class BaseWorkspace(ABC):
    """Workspace 基类，负责文档会话管理"""

    @abstractmethod
    async def execute(self, action: OfficeAction) -> OfficeObs:
        """执行统一动作接口"""
        pass

    @abstractmethod
    def get_document_status(self, document_uri: str) -> DocumentStatus:
        """获取文档状态 (是否连接、是否活跃等)"""
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

### 2.3 职责分层

| 层级 | 职责 | 验证 |
|------|------|------|
| **MCP Tool** | 参数验证 (Pydantic Schema) | 参数格式、必填字段 |
| **Workspace** | 业务验证 + 状态检查 | 文档是否打开、是否可操作 |
| **Socket.IO** | 事件路由 | 连接是否存在 |

---

## 3. 目录结构

```
office4ai/
├── a2c_smcp/                           # MCP 基础设施
│   ├── server.py                       # BaseMCPServer
│   ├── config.py                       # MCPServerConfig
│   ├── tools/                          # MCP 工具 (未来添加)
│   │   ├── base.py                     # BaseTool 基类
│   │   ├── word/                       # Word 工具 (未来)
│   │   ├── ppt/                        # PPT 工具 (未来)
│   │   └── excel/                      # Excel 工具 (未来)
│   └── resources/                      # MCP 资源
│
├── environment/
│   └── workspace/                      # Workspace 核心实现
│       ├── __init__.py
│       ├── base.py                     # ⭐ BaseWorkspace 抽象基类
│       ├── office_workspace.py         # ⭐ OfficeWorkspace 实现
│       │
│       ├── socketio/                   # Workspace Socket.IO Server
│       │   ├── server.py               # Socket.IO 服务器入口
│       │   ├── namespaces/             # 命名空间实现
│       │   │   ├── base.py             # BaseNamespace
│       │   │   ├── word.py             # /word namespace
│       │   │   ├── ppt.py              # /ppt namespace (未来)
│       │   │   └── excel.py            # /excel namespace (未来)
│       │   ├── services/               # 控制服务
│       │   │   └── connection_manager.py  # 连接管理器
│       │   └── middleware/             # 中间件
│       │       └── handshake.py        # 握手中间件
│       │
│       └── dtos/                       # 数据传输对象 (与 TS 严格同步)
│           ├── common.py               # 通用类型
│           ├── word.py                 # Word 类型 (13 个事件)
│           ├── ppt.py                  # PPT 类型 (10 个事件)
│           └── excel.py                # Excel 类型 (4 个事件)
│
└── office/
    └── mcp/
        └── server.py                   # OfficeMCPServer (未来)
```

---

## 4. 实现阶段

### 阶段 1：核心基类 (当前优先)

1. **`environment/workspace/base.py`**
   - `BaseWorkspace` 抽象基类
   - `OfficeAction` / `OfficeObs` 数据模型
   - `DocumentStatus` 状态枚举

2. **`environment/workspace/office_workspace.py`**
   - `OfficeWorkspace` 实现
   - 集成 `ConnectionManager`
   - 实现状态验证 (混合模式：主动探测 + 被动监听)

### 阶段 2：Word 示例 (MVP)

1. 完善 `/word` namespace 的 3 个核心事件：
   - `word:get:selectedContent` (获取选中内容)
   - `word:insert:text` (插入文本)
   - `word:replace:selection` (替换选中内容)

2. 实现端到端调用链路：
   - 测试脚本 → Workspace → Socket.IO → Add-In

### 阶段 3：MCP Tools (后续)

1. 实现 `BaseTool` 基类
2. 实现 Word 工具 (映射到 27 个工具列表，见 `mcp_tools_list.md`)
3. 实现工具注册机制

---

## 5. 文档定位机制

采用 **URI 映射** 方式：

```python
# Workspace 维护 document_uri → socket_id 的映射
class OfficeWorkspace:
    def __init__(self):
        self._document_map: dict[str, str] = {}  # uri → socket_id

    async def execute(self, action: OfficeAction) -> OfficeObs:
        document_uri = action.params["document_uri"]
        socket_id = self._document_map.get(document_uri)

        if not socket_id:
            raise DocumentNotConnected(document_uri)

        # 发送 Socket.IO 事件
        event = f"{action.category}:{action.action_name}"
        return await self.emit_to_document(document_uri, event, action.params)
```

---

## 6. 类型同步策略

Python Pydantic 模型与 TypeScript 类型 **严格同步**：

| Python 文件 | TypeScript 文件 |
|------------|----------------|
| `dtos/common.py` | `shared/socketio-types.ts` (BaseRequest/BaseResponse) |
| `dtos/word.py` | `shared/socketio-types.ts` (WordEvents) |
| `dtos/ppt.py` | `shared/socketio-types.ts` (PptEvents) |
| `dtos/excel.py` | `shared/socketio-types.ts` (ExcelEvents) |

> 当任一侧修改时，必须同步更新另一侧，确保 Socket.IO 事件兼容。

---

## 7. 状态验证实现

采用 **混合模式**：

1. **被动监听**：监听 Socket.IO 的 `disconnect` 事件
2. **主动探测**：定期发送 `ping` 事件检查连接
3. **状态缓存**：维护文档状态表，快速查询

```python
class DocumentStatus(Enum):
    CONNECTED = "connected"       # 已连接且活跃
    DISCONNECTED = "disconnected" # 已断开
    UNKNOWN = "unknown"           # 未知 (需要探测)
```

---

## 8. 参考

- [MCP 工具列表](mcp_tools_list.md) - 27 个工具详细定义
- [A2C 协议](a2c/a2c_rfc.md) - Agent-Computer 协议规范 (独立系统)
- [Office Add-In Socket.IO API](https://github.com/your-org/office-editor4ai) - TypeScript 实现
