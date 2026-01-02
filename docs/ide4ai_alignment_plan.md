# Office4AI A2C 协议对齐实现规划

**文档版本**: v1.0
**创建日期**: 2026-01-02
**参考基准**: `examples/ide4ai` 项目架构
**目标**: 对齐 ide4ai 的大多数功能，实现 Office4AI 的 A2C 协议架构

---

## 1. 执行摘要

本规划基于 `examples/ide4ai` 的成熟架构，为 Office4AI 提供了一个清晰的实施路径。核心目标是：

- **架构对齐**: 实现与 ide4ai 相同的层次化架构模式
- **功能对齐**: 实现 ide4ai 的核心机制，包括环境抽象、工具系统、资源管理
- **差异化实现**: 针对 Office 文档的特殊性，设计专门的文档环境（而非代码 IDE）

---

## 2. ide4ai 架构深度分析

### 2.1 核心架构组件

```
ide4ai/
├── base.py                    # IDE 基类 (gym.Env + ABC)
├── schema.py                  # 动作/观察 Schema 定义
├── ides.py                    # 单例模式管理
├── environment/               # 环境组件
│   ├── workspace/            # 工作区环境 (文件操作 + LSP)
│   └── terminal/             # 终端环境
├── a2c_smcp/                  # MCP 协议层
│   ├── server.py             # BaseMCPServer 基类
│   ├── tools/                # 工具实现
│   │   ├── base.py           # BaseTool 基类
│   │   ├── read.py           # Read 工具
│   │   ├── edit.py           # Edit 工具
│   │   ├── glob.py           # Glob 工具
│   │   ├── grep.py           # Grep 工具
│   │   ├── bash.py           # Bash 工具
│   │   └── write.py          # Write 工具
│   └── resources/            # 资源实现
│       └── base.py           # BaseResource 基类
├── dtos/                      # 数据传输对象
└── python_ide/               # Python IDE 具体实现
```

### 2.2 关键设计模式

#### 2.2.1 双层抽象模式

```
BaseMCPServer (MCP 协议层)
    ↓ 依赖
IDE (强化学习环境层 - gym.Env)
    ↓ 组合
Workspace + Terminal (具体能力层)
```

**关键点**:
- MCP Server 不直接操作文件，而是通过 IDE 的 step() 方法
- IDE 实现了 gymnasium.Env 接口，支持 RL 训练
- Workspace 和 Terminal 都是独立的 gym.Env

#### 2.2.2 工具模式

```python
class BaseTool(ABC):
    def __init__(self, ide: IDE):
        self.ide = ide

    @property
    @abstractmethod
    def name(self) -> str: ...

    @property
    @abstractmethod
    def description(self) -> str: ...

    @property
    @abstractmethod
    def input_schema(self) -> dict[str, Any]: ...

    @abstractmethod
    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]: ...

    def validate_input(self, arguments: dict, model: Type[T]) -> T:
        return model.model_validate(arguments)
```

**关键点**:
- 工具持有 IDE 实例的引用
- 工具通过 `self.ide.workspace` 或 `self.ide.terminal` 调用能力
- 使用 Pydantic 进行参数验证

#### 2.2.3 Schema 驱动模式

```python
class IDEAction(BaseModel):
    category: Literal["terminal", "workspace"]
    action_name: str
    action_args: dict | Json[Any] | list[str] | str | None

class IDEObs(BaseModel):
    created_at: str
    obs: str | None
    original_result: Any
```

**关键点**:
- 使用 Pydantic 定义所有数据结构
- 支持多种参数格式（dict/Json/list/str）
- 统一的观察返回格式

### 2.3 ide4ai 的核心能力

1. **文件操作能力** (Workspace)
   - `open_file()` - 打开文件
   - `read_file()` - 读取文件
   - `apply_edit()` - 应用编辑
   - `save_file()` - 保存文件
   - `create_file()` - 创建文件
   - `delete_file()` - 删除文件
   - `rename_file()` - 重命名文件

2. **搜索能力** (Workspace)
   - `find_in_path()` - 在文件/文件夹中搜索
   - `replace_in_file()` - 在文件中替换
   - `glob_files()` - 使用通配符匹配文件

3. **LSP 能力** (Workspace)
   - `send_lsp_msg()` - 发送 LSP 消息
   - `pull_diagnostics()` - 拉取诊断信息
   - `get_file_symbols()` - 获取文件符号

---

## 3. Office4AI 当前架构对比

### 3.1 已有组件 ✅

```
office4ai/
├── a2c_smcp/
│   ├── server.py              # BaseMCPServer ✅ (已完成)
│   ├── config.py              # MCPServerConfig ✅ (已完成)
│   ├── tools/
│   │   └── base.py            # BaseTool ✅ (简化版)
│   └── resources/
│       └── base.py            # BaseResource ✅ (简化版)
└── office/
    └── mcp/
        └── server.py          # OfficeMCPServer ✅ (已存在但未实现)
```

### 3.2 缺失组件 ❌

| 组件 | ide4ai 位置 | office4ai 状态 | 优先级 |
|------|-------------|----------------|--------|
| **IDE 基类** | `base.py: IDE` | ❌ 缺失 | **P0** |
| **Schema 定义** | `schema.py` | ❌ 缺失 | **P0** |
| **文档环境基类** | `environment/workspace/base.py: BaseWorkspace` | ❌ 缺失 | **P0** |
| **具体工具实现** | `a2c_smcp/tools/` | ❌ 缺失 | **P1** |
| **Schema/DTO** | `dtos/` | ❌ 缺失 | **P1** |
| **资源实现** | `a2c_smcp/resources/` | ❌ 缺失 | **P2** |
| **单例管理** | `ides.py` | ❌ 缺失 | **P3** |

### 3.3 差异分析

#### 架构层面
- ✅ **MCP 层**: 已有 BaseMCPServer，与 ide4ai 对齐
- ❌ **环境层**: 缺少 Office IDE 基类和文档环境基类
- ❌ **工具层**: BaseTool 过于简化，缺少 validate_input 等辅助方法
- ❌ **Schema 层**: 完全缺失

#### 功能层面
- ✅ **传输协议**: 已支持 stdio/sse/streamable-http
- ❌ **工具注册**: OfficeMCPServer 未实现 _register_tools/_register_resources
- ❌ **文档操作**: 没有对应的文档环境实现
- ❌ **资源管理**: 没有实现任何资源

---

## 4. 实施路线图

### Phase 1: 基础架构对齐 (Week 1-2)

**目标**: 建立与 ide4ai 对齐的基础架构

#### 1.1 Schema 定义 (P0)

**文件**: `office4ai/office/schema.py`

```python
from pydantic import BaseModel, Field
from typing import Any, Literal, Union
from datetime import datetime

class OfficeObs(BaseModel):
    """Office 环境观察"""
    created_at: str = Field(
        default_factory=lambda: datetime.now().isoformat()
    )
    obs: str | None = None
    original_result: Any = None

class OfficeAction(BaseModel):
    """Office 环境动作"""
    category: Literal["document", "workspace", "terminal"]
    action_name: str
    action_args: dict | list[str] | str | None = None

class DocumentType(str, Enum):
    """文档类型"""
    WORD = "writer"  # Writer 文档
    CALC = "calc"    # Calc 表格
    IMPRESS = "impress"  # Impress 演示
```

#### 1.2 文档环境基类 (P0)

**文件**: `office4ai/office/environment/base.py`

```python
from abc import ABC, abstractmethod
from gymnasium import Env
from office4ai.office.schema import OfficeAction, OfficeObs

class BaseDocumentEnv(Env, ABC):
    """
    文档环境基类

    对齐 ide4ai 的 BaseWorkspace，但针对 Office 文档
    """

    @abstractmethod
    def open_document(self, uri: str) -> Any:
        """打开文档"""
        pass

    @abstractmethod
    def read_document(self, uri: str) -> str:
        """读取文档内容"""
        pass

    @abstractmethod
    def apply_edit(self, uri: str, edits: list) -> tuple:
        """应用编辑"""
        pass

    @abstractmethod
    def save_document(self, uri: str) -> None:
        """保存文档"""
        pass

    @abstractmethod
    def find_in_document(
        self,
        uri: str,
        query: str,
        is_regex: bool = False,
        match_case: bool = False
    ) -> list:
        """在文档中搜索"""
        pass
```

#### 1.3 Office IDE 基类 (P0)

**文件**: `office4ai/office/base.py`

```python
from gymnasium import Env
from office4ai.office.environment.base import BaseDocumentEnv

class OfficeIDE(Env, ABC):
    """
    Office IDE 基类

    对齐 ide4ai 的 IDE 基类
    """

    def __init__(
        self,
        workspace_dir: str,
        project_name: str,
        **kwargs
    ):
        super().__init__(**kwargs)
        self.workspace_dir = workspace_dir
        self.project_name = project_name
        self.documents: list[BaseDocumentEnv] = []
        self.active_document_index: int | None = None

    @abstractmethod
    def init_document_env(self) -> BaseDocumentEnv:
        """初始化文档环境"""
        pass

    @property
    def document(self) -> BaseDocumentEnv:
        """获取当前活动文档"""
        if self.active_document_index is not None:
            return self.documents[self.active_document_index]
        else:
            doc = self.init_document_env()
            self.documents.append(doc)
            self.active_document_index = len(self.documents) - 1
            return doc

    def step(self, action: dict) -> tuple:
        """执行一步"""
        office_action = OfficeAction.model_validate(action)
        # 路由到对应的环境
        if office_action.category == "document":
            return self.document.step(action)
        # ... 其他类别
```

**验收标准**:
- [ ] Schema 定义完成，包含所有必要字段
- [ ] BaseDocumentEnv 抽象类定义完整
- [ ] OfficeIDE 基类可以实例化（不需要具体实现）
- [ ] 所有抽象方法都有清晰的文档说明

---

### Phase 2: 工具系统完善 (Week 3)

**目标**: 实现与 ide4ai 对齐的工具系统

#### 2.1 增强 BaseTool (P1)

**文件**: `office4ai/a2c_smcp/tools/base.py`

```python
from abc import ABC, abstractmethod
from typing import Any, TypeVar
from pydantic import BaseModel

T = TypeVar("T", bound=BaseModel)

class BaseTool(ABC):
    """工具基类 - 对齐 ide4ai"""

    def __init__(self, office_ide):
        self.office_ide = office_ide

    @property
    @abstractmethod
    def name(self) -> str:
        """工具名称"""
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """工具描述"""
        pass

    @property
    @abstractmethod
    def input_schema(self) -> dict[str, Any]:
        """输入 Schema (JSON Schema)"""
        pass

    @abstractmethod
    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """执行工具"""
        pass

    def validate_input(self, arguments: dict, model: type[T]) -> T:
        """验证输入参数"""
        try:
            return model.model_validate(arguments)
        except Exception as e:
            raise ValueError(f"参数验证失败: {e}") from e
```

#### 2.2 实现核心工具 (P1)

**优先工具列表** (按优先级排序):

1. **ReadDocument** - 读取文档内容
   - 文件: `office4ai/a2c_smcp/tools/read.py`
   - 参考: `ide4ai/a2c_smcp/tools/read.py`

2. **EditDocument** - 编辑文档
   - 文件: `office4ai/a2c_smcp/tools/edit.py`
   - 参考: `ide4ai/a2c_smcp/tools/edit.py`

3. **GlobDocuments** - 查找文档
   - 文件: `office4ai/a2c_smcp/tools/glob.py`
   - 参考: `ide4ai/a2c_smcp/tools/glob.py`

4. **FindInDocument** - 在文档中搜索
   - 文件: `office4ai/a2c_smcp/tools/find.py`
   - 参考: `ide4ai/a2c_smcp/tools/grep.py`

5. **WriteDocument** - 写入文档
   - 文件: `office4ai/a2c_smcp/tools/write.py`
   - 参考: `ide4ai/a2c_smcp/tools/write.py`

**每个工具的结构**:

```python
from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.a2c_smcp.schemas import ReadInput, ReadOutput

class ReadDocumentTool(BaseTool):
    """文档读取工具"""

    @property
    def name(self) -> str:
        return "ReadDocument"

    @property
    def description(self) -> str:
        return """从 Office 文档中读取内容。

使用说明：
- 支持 Writer、Calc、Impress 文档
- 可以读取文档的文本内容、结构、样式等
- 支持按页、按节、按段落读取
"""

    @property
    def input_schema(self) -> dict[str, Any]:
        return ReadInput.model_json_schema()

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        # 1. 验证输入
        try:
            read_input = self.validate_input(arguments, ReadInput)
        except ValueError as e:
            return ReadOutput(
                success=False,
                content="",
                error=str(e)
            ).model_dump()

        # 2. 检查环境
        if not self.office_ide.document:
            raise RuntimeError("文档环境未初始化")

        # 3. 执行操作
        content = self.office_ide.document.read_document(
            uri=read_input.uri,
            options=read_input.options
        )

        # 4. 返回结果
        return ReadOutput(
            success=True,
            content=content
        ).model_dump()
```

**验收标准**:
- [ ] BaseTool 包含所有必要的方法和辅助函数
- [ ] 至少实现 3 个核心工具 (Read/Edit/Glob)
- [ ] 每个工具都有完整的 Schema 定义
- [ ] 工具可以通过 MCP 调用并返回正确结果

---

### Phase 3: Schema/DTO 系统 (Week 3-4)

**目标**: 建立完整的 Schema 系统

#### 3.1 工具 Schema (P1)

**文件**: `office4ai/a2c_smcp/schemas/__init__.py`

```python
from pydantic import BaseModel, Field
from typing import Optional, Literal

# ========== Read Tool ==========
class ReadInput(BaseModel):
    """Read 工具输入"""
    uri: str = Field(..., description="文档 URI")
    page_index: Optional[int] = Field(None, description="页索引 (从 0 开始)")
    section_index: Optional[int] = Field(None, description="节索引")
    with_structure: bool = Field(False, description="是否包含结构信息")

class ReadOutput(BaseModel):
    """Read 工具输出"""
    success: bool
    content: str
    metadata: Optional[dict] = None
    error: Optional[str] = None

# ========== Edit Tool ==========
class EditInput(BaseModel):
    """Edit 工具输入"""
    uri: str = Field(..., description="文档 URI")
    old_text: str = Field(..., description="要替换的文本")
    new_text: str = Field(..., description="新文本")
    replace_all: bool = Field(False, description="是否替换所有匹配项")

class EditOutput(BaseModel):
    """Edit 工具输出"""
    success: bool
    message: str
    replacements_made: int
    metadata: Optional[dict] = None
    error: Optional[str] = None

# ========== Glob Tool ==========
class GlobInput(BaseModel):
    """Glob 工具输入"""
    pattern: str = Field(..., description="通配符模式")
    path: Optional[str] = Field(None, description="搜索路径")

class GlobOutput(BaseModel):
    """Glob 工具输出"""
    success: bool
    files: list[dict]
    total_count: int
    error: Optional[str] = None
```

#### 3.2 环境 Schema (P1)

**文件**: `office4ai/office/environment/schema.py`

```python
from pydantic import BaseModel
from typing import Optional, Literal

class Position(BaseModel):
    """位置"""
    line: int
    character: int

class Range(BaseModel):
    """范围"""
    start: Position
    end: Position

class SearchResult(BaseModel):
    """搜索结果"""
    range: Range
    matched_text: str
    context_before: str
    context_after: str

class DocumentMetadata(BaseModel):
    """文档元数据"""
    uri: str
    document_type: Literal["writer", "calc", "impress"]
    page_count: int
    title: str
    author: Optional[str]
    created_at: str
    modified_at: str
```

**验收标准**:
- [ ] 所有工具都有对应的 Input/Output Schema
- [ ] Schema 包含详细的 Field 描述和验证规则
- [ ] 环境相关的 Schema 定义完整
- [ ] 所有 Schema 都通过 Pydantic 验证测试

---

### Phase 4: OfficeMCPServer 实现 (Week 4)

**目标**: 实现完整的 OfficeMCPServer

#### 4.1 OfficeMCPServer 主类 (P1)

**文件**: `office4ai/office/mcp/server.py`

```python
from office4ai.a2c_smcp.server import BaseMCPServer
from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.office.base import OfficeIDE
from office4ai.a2c_smcp.tools import (
    ReadDocumentTool,
    EditDocumentTool,
    GlobDocumentsTool,
    FindInDocumentTool,
)

class OfficeMCPServer(BaseMCPServer):
    """Office MCP Server 实现"""

    def __init__(self, config: MCPServerConfig):
        super().__init__(config, "office4ai")

        # 创建 Office IDE 实例
        self.office_ide = self._create_office_ide()

    def _create_office_ide(self) -> OfficeIDE:
        """创建 Office IDE 实例"""
        # 这里暂时返回 None，具体实现在后续阶段
        # TODO: 实现具体的 Office IDE (LibreOffice/UNO)
        return None

    def _register_tools(self) -> None:
        """注册所有工具"""
        tools = [
            ReadDocumentTool(self.office_ide),
            EditDocumentTool(self.office_ide),
            GlobDocumentsTool(self.office_ide),
            FindInDocumentTool(self.office_ide),
        ]

        for tool in tools:
            self.tools[tool.name] = tool

    def _register_resources(self) -> None:
        """注册所有资源"""
        # 暂时不实现资源，在 Phase 5 完成
        pass
```

**验收标准**:
- [ ] OfficeMCPServer 可以成功启动
- [ ] `list_tools` 返回已注册的工具列表
- [ ] `call_tool` 可以正确调用工具
- [ ] 错误处理和日志记录正常工作

---

### Phase 5: 资源系统 (Week 5)

**目标**: 实现资源管理功能

#### 5.1 增强基类 (P2)

**文件**: `office4ai/a2c_smcp/resources/base.py`

```python
from abc import ABC, abstractmethod
from typing import Optional
from urllib.parse import urlparse, parse_qs

class BaseResource(ABC):
    """资源基类 - 对齐 ide4ai"""

    def __init__(self, base_uri: str, office_ide):
        self.base_uri = base_uri
        self.uri = base_uri
        self.office_ide = office_ide
        self._params = {}

    @property
    @abstractmethod
    def name(self) -> str:
        """资源名称"""
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """资源描述"""
        pass

    @property
    @abstractmethod
    def mime_type(self) -> str:
        """MIME 类型"""
        pass

    @abstractmethod
    async def read(self) -> str:
        """读取资源内容"""
        pass

    def update_from_uri(self, uri: str) -> None:
        """从 URI 更新参数"""
        parsed = urlparse(uri)
        self._params = parse_qs(parsed.query)
        self.uri = uri
```

#### 5.2 实现核心资源 (P2)

1. **ActiveDocumentResource** - 当前活动文档
   - URI: `office://document/active`
   - 返回: 文档元信息 (JSON)

2. **DocumentOutlineResource** - 文档大纲
   - URI: `office://document/outline`
   - 返回: 标题层级结构 (JSON)

3. **DocumentPageResource** - 文档页面
   - URI: `office://document/page?index=1`
   - 返回: 渲染的页面图像 (base64)

**验收标准**:
- [ ] BaseResource 实现完整
- [ ] 至少实现 2 个核心资源
- [ ] `list_resources` 返回已注册的资源列表
- [ ] `read_resource` 可以正确读取资源

---

### Phase 6: 文档环境实现 (Week 6-8)

**目标**: 实现具体的文档环境（依赖 Office Add-In）

**注意**: 这是用户提到的"难点但非重点"，这里只做架构封装，不具体实现

#### 6.1 LibreOffice 环境基类 (P2)

**文件**: `office4ai/office/environment/libreoffice.py`

```python
from office4ai.office.environment.base import BaseDocumentEnv

class LibreOfficeEnv(BaseDocumentEnv):
    """
    LibreOffice 文档环境

    注意: 这是架构封装，具体实现需要与 Office Add-In 打通
    """

    def __init__(self, uno_bridge):
        self.uno_bridge = uno_bridge  # 由 Office Add-In 提供

    def open_document(self, uri: str) -> Any:
        """
        打开文档

        实际实现:
        1. 调用 Office Add-In 的 API
        2. 获取文档句柄
        3. 返回文档对象
        """
        # TODO: 调用 Office Add-In API
        raise NotImplementedError("需要与 Office Add-In 打通")

    def read_document(self, uri: str) -> str:
        """读取文档内容"""
        # TODO: 调用 Office Add-In API
        raise NotImplementedError("需要与 Office Add-In 打通")

    # ... 其他方法
```

#### 6.2 Office Add-In 接口定义 (P2)

**文件**: `office4ai/office/addin/interface.py`

```python
from abc import ABC, abstractmethod

class OfficeAddInInterface(ABC):
    """
    Office Add-In 接口定义

    定义所有需要 Office Add-In 实现的接口
    """

    @abstractmethod
    def open_document(self, file_path: str) -> str:
        """
        打开文档

        Returns:
            文档 ID
        """
        pass

    @abstractmethod
    def get_document_text(self, doc_id: str) -> str:
        """获取文档文本"""
        pass

    @abstractmethod
    def replace_text(
        self,
        doc_id: str,
        old_text: str,
        new_text: str,
        replace_all: bool = False
    ) -> int:
        """
        替换文本

        Returns:
            替换次数
        """
        pass

    @abstractmethod
    def find_text(
        self,
        doc_id: str,
        query: str,
        match_case: bool = False
    ) -> list[dict]:
        """查找文本"""
        pass

    # ... 更多接口
```

**验收标准**:
- [ ] LibreOfficeEnv 基类定义完整
- [ ] OfficeAddInInterface 接口定义清晰
- [ ] 所有方法都有文档说明
- [ ] 提供一个 Mock 实现用于测试

---

### Phase 7: 单例管理与生命周期 (Week 9)

**目标**: 实现单例模式和资源管理

#### 7.1 单例管理 (P3)

**文件**: `office4ai/office/singleton.py`

```python
import threading
from typing import Any

class OfficeIDESingleton(type):
    """Office IDE 单例元类"""
    _instances = {}
    _lock = threading.Lock()

    def __call__(cls, *args, **kwargs):
        key = cls.__name__ + kwargs.get("project_name", "")
        if key not in cls._instances:
            with cls._lock:
                if key not in cls._instances:
                    cls._instances[key] = super().__call__(*args, **kwargs)
        return cls._instances[key]

class OfficeIDEInstance(metaclass=OfficeIDESingleton):
    """Office IDE 单例实例"""

    def __init__(
        self,
        workspace_dir: str,
        project_name: str,
        **kwargs
    ):
        from office4ai.office.base import OfficeIDE

        self._ide = OfficeIDE(
            workspace_dir=workspace_dir,
            project_name=project_name,
            **kwargs
        )

    @property
    def ide(self):
        return self._ide
```

#### 7.2 生命周期管理 (P2)

**关键点**:
- 实现 `close()` 方法，确保资源正确释放
- 在 `__del__` 中调用 `close()`
- 使用 `atexit` 注册清理函数

**验收标准**:
- [ ] 单例模式正常工作
- [ ] 资源正确释放（无内存泄漏）
- [ ] 异常情况下也能正确清理

---

## 5. 优先级矩阵

| 任务 | 优先级 | 复杂度 | 依赖 | 周期 |
|------|--------|--------|------|------|
| Schema 定义 | P0 | 低 | 无 | 1-2 天 |
| 基类定义 | P0 | 中 | Schema | 2-3 天 |
| BaseTool 增强 | P1 | 低 | 无 | 1 天 |
| 核心工具实现 (3个) | P1 | 中 | BaseTool | 3-5 天 |
| Schema/DTO 系统 | P1 | 中 | 工具定义 | 2-3 天 |
| OfficeMCPServer 实现 | P1 | 中 | 工具 + 基类 | 2-3 天 |
| 资源系统 | P2 | 中 | BaseMCPServer | 2-3 天 |
| 文档环境架构 | P2 | 高 | 无 | 3-5 天 |
| Office Add-In 接口 | P2 | 中 | 文档环境 | 2-3 天 |
| 单例管理 | P3 | 低 | OfficeIDE | 1-2 天 |

---

## 6. 风险与缓解策略

### 6.1 架构风险

**风险 1**: 过度设计，引入不必要的复杂性
- **缓解**: 严格对齐 ide4ai 架构，不添加额外抽象
- **指标**: 代码行数与 ide4ai 相近

**风险 2**: 抽象过早，难以适应实际需求
- **缓解**: 先实现简化版本，再逐步完善
- **指标**: 每个阶段都有可运行的代码

### 6.2 实现风险

**风险 3**: Office Add-In 接口不匹配
- **缓解**: 在 Phase 6 定义清晰的接口契约
- **指标**: 接口定义完整，有 Mock 实现

**风险 4**: 工具实现不一致
- **缓解**: 严格遵守 ide4ai 的工具模式
- **指标**: 所有工具通过相同的测试套件

### 6.3 集成风险

**风险 5**: MCP 协议兼容性问题
- **缓解**: 使用 ide4ai 的 BaseMCPServer 作为基类
- **指标**: 通过 MCP 标准测试

---

## 7. 测试策略

### 7.1 单元测试

```python
# tests/unit_tests/test_schema.py
def test_office_action_validation():
    action = OfficeAction(
        category="document",
        action_name="read",
        action_args={"uri": "file:///test.docx"}
    )
    assert action.category == "document"

# tests/unit_tests/test_tools.py
async def test_read_document_tool():
    tool = ReadDocumentTool(mock_ide)
    result = await tool.execute({
        "uri": "file:///test.docx"
    })
    assert result["success"] is True
```

### 7.2 集成测试

```python
# tests/integration_tests/test_mcp_server.py
async def test_mcp_server_startup():
    config = MCPServerConfig(transport="stdio")
    server = OfficeMCPServer(config)
    # 测试服务器启动和工具注册
```

### 7.3 端到端测试

```python
# tests/e2e_tests/test_document_workflow.py
async def test_complete_document_workflow():
    """测试完整的文档工作流"""
    # 1. 创建文档
    # 2. 添加内容
    # 3. 保存文档
    # 4. 导出 PDF
```

---

## 8. 文档清单

- [ ] 架构设计文档 (`docs/architecture.md`)
- [ ] API 文档 (`docs/api.md`)
- [ ] 工具使用指南 (`docs/tools.md`)
- [ ] 资源使用指南 (`docs/resources.md`)
- [ ] 集成指南 (`docs/integration.md`)
- [ ] 故障排查指南 (`docs/troubleshooting.md`)

---

## 9. 交付物清单

### Phase 1 交付物
- [ ] `office4ai/office/schema.py` - Schema 定义
- [ ] `office4ai/office/environment/base.py` - 环境基类
- [ ] `office4ai/office/base.py` - Office IDE 基类

### Phase 2 交付物
- [ ] `office4ai/a2c_smcp/tools/base.py` - 增强的 BaseTool
- [ ] `office4ai/a2c_smcp/tools/read.py` - Read 工具
- [ ] `office4ai/a2c_smcp/tools/edit.py` - Edit 工具
- [ ] `office4ai/a2c_smcp/tools/glob.py` - Glob 工具

### Phase 3 交付物
- [ ] `office4ai/a2c_smcp/schemas/__init__.py` - 工具 Schema
- [ ] `office4ai/office/environment/schema.py` - 环境 Schema

### Phase 4 交付物
- [ ] `office4ai/office/mcp/server.py` - 完整的 OfficeMCPServer

### Phase 5 交付物
- [ ] `office4ai/a2c_smcp/resources/base.py` - 资源基类
- [ ] `office4ai/a2c_smcp/resources/document.py` - 文档资源

### Phase 6 交付物
- [ ] `office4ai/office/environment/libreoffice.py` - LibreOffice 环境
- [ ] `office4ai/office/addin/interface.py` - Office Add-In 接口

### Phase 7 交付物
- [ ] `office4ai/office/singleton.py` - 单例管理

---

## 10. 下一步行动

### 立即开始 (本周)
1. 创建 `office4ai/office/schema.py`
2. 创建 `office4ai/office/environment/base.py`
3. 创建 `office4ai/office/base.py`

### 短期目标 (2周内)
1. 完成 Phase 1-2
2. 实现至少 3 个核心工具
3. 建立 CI/CD pipeline

### 中期目标 (1-2月)
1. 完成 Phase 3-5
2. 实现资源系统
3. 完成单元测试覆盖

### 长期目标 (3-6月)
1. 完成 Phase 6-7
2. 与 Office Add-In 打通
3. 发布 v1.0 版本

---

## 11. 附录

### A. ide4ai 关键代码片段

#### A.1 IDE.step() 实现
```python
def step(self, action: dict) -> tuple:
    ide_action = self.construct_action(action)
    if ide_action.category == "terminal":
        return self.terminal.step(action)
    else:
        if self.workspace:
            return self.workspace.step(action)
        else:
            raise IDEExecutionError("Workspace 未初始化")
```

#### A.2 Workspace.read_file() 实现
```python
def read_file(self, *, uri: str, with_line_num: bool = True) -> str:
    tm = next(filter(lambda m: m.uri == AnyUrl(uri), self.models), None)
    if tm:
        return tm.get_view(with_line_num)
    else:
        tm = self.open_file(uri=uri)
        return tm.get_view(with_line_num)
```

### B. 参考资料

- [ide4ai GitHub 仓库](https://github.com/your-org/ide4ai)
- [MCP 协议规范](https://modelcontextprotocol.io)
- [Gymnasium 文档](https://gymnasium.farama.org)
- [Pydantic 文档](https://docs.pydantic.dev)

---

**文档结束**
