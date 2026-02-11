# Office4AI CI/CD 自动化测试架构规划

> **版本**: v1.0.0  
> **创建日期**: 2026-01-08  
> **状态**: 待实施  
> **作者**: JQQ

---

## 1. 背景与目标

### 1.1 现状问题

当前 E2E 测试（`manual_tests/`）需要：
- 手动启动 Workspace 服务器
- 手动启动 Word Add-In 开发服务器
- 手动打开 Office 桌面应用并加载 Add-In
- 人工触发操作并验证结果

这种方式无法在 CI/CD 中自动运行，无法保证代码质量。

### 1.2 目标

建立可在 CI/CD 中自动运行的测试体系，覆盖：
- Python 端所有代码（Workspace、DTO、Service、Namespace）
- 模拟 Add-In 响应行为（Mock Add-In Client）
- 完整的请求-响应流程验证

### 1.3 边界定义

| 范围 | 本方案覆盖 | 说明 |
|------|-----------|------|
| Python Workspace 服务器 | ✅ | 完整覆盖 |
| Socket.IO 通信协议 | ✅ | 通过 Mock Client 验证 |
| Add-In 响应行为模拟 | ✅ | Mock Add-In Client |
| TypeScript Add-In 代码 | ❌ | 由 office-editor4ai 项目独立测试 |
| 真实 Office 应用交互 | ❌ | 保留在 manual_tests/ |

---

## 2. 测试分层架构

```
┌─────────────────────────────────────────────────────────────────┐
│                        测试金字塔                                │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│   ┌─────────────────┐                                          │
│   │  Manual Tests   │  ← 人工验证（保留现有 manual_tests/）      │
│   │   (E2E Real)    │    真实 Office + Add-In                   │
│   └────────┬────────┘                                          │
│            │                                                    │
│   ┌────────▼────────┐                                          │
│   │ Contract Tests  │  ← 协议契约测试（新增）                    │
│   │ (Mock Add-In)   │    验证请求/响应格式符合协议               │
│   └────────┬────────┘                                          │
│            │                                                    │
│   ┌────────▼────────┐                                          │
│   │Integration Tests│  ← 集成测试（增强现有）                    │
│   │(Server + Client)│    Socket.IO 服务器 + Mock 客户端         │
│   └────────┬────────┘                                          │
│            │                                                    │
│   ┌────────▼────────┐                                          │
│   │   Unit Tests    │  ← 单元测试（现有）                        │
│   │ (DTO/Service)   │    DTO 验证、Service 逻辑                 │
│   └─────────────────┘                                          │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 2.1 各层职责

| 层级 | 目录 | 职责 | CI/CD |
|------|------|------|-------|
| **Unit Tests** | `tests/unit_tests/` | DTO 验证、Service 业务逻辑、工具函数 | ✅ |
| **Integration Tests** | `tests/integration_tests/` | Socket.IO 服务器生命周期、客户端连接、事件路由 | ✅ |
| **Contract Tests** | `tests/contract_tests/` | 请求/响应协议验证、Mock Add-In 响应 | ✅ |
| **Manual Tests** | `manual_tests/` | 真实 E2E 验证、人工回归测试 | ❌ |

---

## 3. Mock Add-In Client 设计

### 3.1 核心概念

Mock Add-In Client 模拟真实 Add-In 的行为：
1. 连接到 Workspace Socket.IO 服务器
2. 完成握手（发送 clientId、documentUri）
3. 接收服务器发送的事件
4. 返回符合协议的响应

### 3.2 架构设计

```
┌─────────────────────────────────────────────────────────────────┐
│                     Test Environment                            │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌──────────────────┐         ┌──────────────────────────────┐ │
│  │                  │         │                              │ │
│  │  OfficeWorkspace │◄───────►│     MockAddInClient          │ │
│  │  (Real Server)   │ Socket  │     (Test Double)            │ │
│  │                  │   IO    │                              │ │
│  └──────────────────┘         └──────────────────────────────┘ │
│           │                              │                      │
│           │                              │                      │
│  ┌────────▼─────────┐         ┌──────────▼───────────────────┐ │
│  │ ConnectionManager│         │   ResponseRegistry           │ │
│  │ (Real)           │         │   (Configurable Responses)   │ │
│  └──────────────────┘         └──────────────────────────────┘ │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 3.3 MockAddInClient 接口

```python
class MockAddInClient:
    """模拟 Office Add-In 客户端"""
    
    def __init__(
        self,
        server_url: str,
        namespace: str,
        client_id: str,
        document_uri: str,
    ) -> None: ...
    
    async def connect(self) -> None:
        """连接到服务器并完成握手"""
        ...
    
    async def disconnect(self) -> None:
        """断开连接"""
        ...
    
    def register_response(
        self,
        event: str,
        response_factory: Callable[[dict], dict],
    ) -> None:
        """注册事件响应工厂函数"""
        ...
    
    def register_static_response(
        self,
        event: str,
        response: dict,
    ) -> None:
        """注册静态响应"""
        ...
    
    @property
    def received_events(self) -> list[tuple[str, dict]]:
        """获取收到的所有事件（用于断言）"""
        ...
```

### 3.4 响应注册模式

```python
# 模式 1: 静态响应
mock_client.register_static_response(
    "word:get:selectedContent",
    {
        "success": True,
        "data": {
            "text": "Hello World",
            "elements": [],
            "metadata": {"characterCount": 11}
        }
    }
)

# 模式 2: 动态响应工厂
def selected_content_response(request: dict) -> dict:
    return {
        "requestId": request["requestId"],
        "success": True,
        "data": generate_mock_content(),
        "timestamp": int(time.time() * 1000),
    }

mock_client.register_response(
    "word:get:selectedContent",
    selected_content_response
)

# 模式 3: 错误响应
mock_client.register_static_response(
    "word:get:selectedContent",
    {
        "success": False,
        "error": {
            "code": "3002",
            "message": "Selection is empty"
        }
    }
)
```

---

## 4. 测试数据生成

### 4.1 Mock 库选择

**推荐**: [Mimesis](https://mimesis.name/) v18+

选择理由：
- 现代化设计，类型安全
- 支持多语言 locale
- 内置丰富的数据提供者
- 活跃维护，Python 3.10+ 支持

**备选**: [Polyfactory](https://github.com/litestar-org/polyfactory)
- 专为 Pydantic 模型设计
- 自动从模型生成测试数据

### 4.2 依赖添加

```toml
# pyproject.toml [dependency-groups.dev]
"mimesis>=18.0.0",
"polyfactory>=2.0.0",
```

### 4.3 数据工厂设计

```python
# tests/factories/word_factories.py

from mimesis import Field, Locale
from mimesis.providers import Text, Datetime

from office4ai.environment.workspace.dtos.word import (
    WordGetSelectedContentRequest,
    GetContentOptions,
)

field = Field(Locale.EN)

class WordDataFactory:
    """Word 事件测试数据工厂"""
    
    @staticmethod
    def selected_content_response(
        text: str | None = None,
        include_elements: bool = True,
    ) -> dict:
        """生成 word:get:selectedContent 响应数据"""
        text = text or field("text.sentence")
        return {
            "text": text,
            "elements": [
                {"type": "paragraph", "content": text}
            ] if include_elements else [],
            "metadata": {
                "characterCount": len(text),
                "paragraphCount": 1,
                "tableCount": 0,
                "imageCount": 0,
            }
        }
    
    @staticmethod
    def insert_text_response(success: bool = True) -> dict:
        """生成 word:insert:text 响应数据"""
        if success:
            return {
                "insertedLength": field("numeric.integer_number", start=1, end=1000),
                "position": {"start": 0, "end": 100},
            }
        return {}
```

### 4.4 Polyfactory 集成（可选）

```python
from polyfactory.factories.pydantic_factory import ModelFactory
from office4ai.environment.workspace.dtos.word import WordGetSelectedContentRequest

class WordGetSelectedContentRequestFactory(ModelFactory):
    __model__ = WordGetSelectedContentRequest
    
    # 自定义字段生成规则
    document_uri = "file:///tmp/test.docx"

# 使用
request = WordGetSelectedContentRequestFactory.build()
```

---

## 5. Contract Tests 设计

### 5.1 目录结构

```
tests/contract_tests/
├── __init__.py
├── conftest.py                    # Contract 测试 fixtures
├── mock_addin/
│   ├── __init__.py
│   ├── client.py                  # MockAddInClient 实现
│   └── response_registry.py       # 响应注册表
├── factories/
│   ├── __init__.py
│   ├── word_factories.py          # Word 数据工厂
│   ├── ppt_factories.py           # PPT 数据工厂
│   └── excel_factories.py         # Excel 数据工厂
└── word/
    ├── __init__.py
    ├── test_get_selected_content.py
    ├── test_insert_text.py
    └── test_replace_selection.py
```

### 5.2 Contract Test 示例

```python
# tests/contract_tests/word/test_get_selected_content.py

import pytest
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.base import OfficeAction

@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_success(
    workspace: OfficeWorkspace,
    mock_addin_client: MockAddInClient,
    word_factory: WordDataFactory,
):
    """测试成功获取选中内容的完整流程"""
    # Arrange: 注册 Mock 响应
    expected_text = "Test selected content"
    mock_addin_client.register_response(
        "word:get:selectedContent",
        lambda req: {
            "requestId": req["requestId"],
            "success": True,
            "data": word_factory.selected_content_response(text=expected_text),
            "timestamp": int(time.time() * 1000),
        }
    )
    
    # Act: 执行动作
    action = OfficeAction(
        category="word",
        action_name="get:selectedContent",
        params={"document_uri": mock_addin_client.document_uri},
    )
    result = await workspace.execute(action)
    
    # Assert: 验证结果
    assert result.success is True
    assert result.data["text"] == expected_text
    assert "metadata" in result.data


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selected_content_empty_selection(
    workspace: OfficeWorkspace,
    mock_addin_client: MockAddInClient,
):
    """测试选区为空的错误处理"""
    # Arrange: 注册错误响应
    mock_addin_client.register_response(
        "word:get:selectedContent",
        lambda req: {
            "requestId": req["requestId"],
            "success": False,
            "error": {
                "code": "3002",
                "message": "Selection is empty",
            },
            "timestamp": int(time.time() * 1000),
        }
    )
    
    # Act
    action = OfficeAction(
        category="word",
        action_name="get:selectedContent",
        params={"document_uri": mock_addin_client.document_uri},
    )
    result = await workspace.execute(action)
    
    # Assert
    assert result.success is False
    assert "3002" in str(result.error) or "empty" in str(result.error).lower()
```

### 5.3 Fixtures 设计

```python
# tests/contract_tests/conftest.py

import pytest
import pytest_asyncio
from tests.contract_tests.mock_addin.client import MockAddInClient
from tests.contract_tests.factories.word_factories import WordDataFactory
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

@pytest_asyncio.fixture
async def workspace():
    """启动测试用 Workspace"""
    ws = OfficeWorkspace(host="127.0.0.1", port=3002)
    await ws.start()
    yield ws
    await ws.stop()

@pytest_asyncio.fixture
async def mock_addin_client(workspace):
    """创建并连接 Mock Add-In 客户端"""
    client = MockAddInClient(
        server_url="http://127.0.0.1:3002",
        namespace="/word",
        client_id="test_client_001",
        document_uri="file:///tmp/contract_test.docx",
    )
    await client.connect()
    yield client
    await client.disconnect()

@pytest.fixture
def word_factory():
    """Word 数据工厂"""
    return WordDataFactory()
```

---

## 6. CI/CD 配置更新

### 6.1 pytest 配置更新

```toml
# pyproject.toml

[tool.pytest.ini_options]
testpaths = ["tests"]
markers = [
    "flaky: mark test as flaky (may need multiple retries)",
    "integration: mark test as integration test",
    "contract: mark test as contract test",  # 新增
]

[tool.poe.tasks]
# 更新测试任务
test = "pytest tests/unit_tests tests/integration_tests tests/contract_tests"
test-unit = "pytest tests/unit_tests"
test-integration = "pytest tests/integration_tests -m integration"
test-contract = "pytest tests/contract_tests -m contract"  # 新增
test-ci = "pytest tests/unit_tests tests/integration_tests tests/contract_tests --tb=short"  # 新增
```

### 6.2 GitHub Actions 更新

```yaml
# .github/workflows/tests.yml

# ... 现有配置 ...

      # 运行测试（更新）
      - name: Run tests
        run: |
          uv run poe test-ci
```

### 6.3 测试执行顺序

CI/CD 中测试执行顺序：
1. **Unit Tests** - 最快，无外部依赖
2. **Integration Tests** - Socket.IO 服务器测试
3. **Contract Tests** - Mock Add-In 完整流程

---

## 7. 实施计划

### Phase 1: 基础设施（1-2 天）

1. 添加依赖：`mimesis`, `polyfactory`
2. 创建 `tests/contract_tests/` 目录结构
3. 实现 `MockAddInClient` 基础版本
4. 更新 pytest 配置

### Phase 2: Word 事件覆盖（2-3 天）

1. 实现 `WordDataFactory`
2. 为已实现的 Word 事件编写 Contract Tests：
   - `word:get:selectedContent`
   - `word:insert:text`
   - `word:replace:selection`
3. 覆盖成功和错误场景

### Phase 3: 扩展与优化（持续）

1. 添加 PPT、Excel 事件覆盖
2. 优化 Mock 响应延迟模拟
3. 添加超时场景测试
4. 集成覆盖率报告

---

## 8. 验收标准

### 8.1 功能验收

- [ ] `MockAddInClient` 能够连接到 Workspace 并完成握手
- [ ] 可以注册静态和动态响应
- [ ] Contract Tests 能够验证完整的请求-响应流程
- [ ] 所有测试在 CI/CD 中自动运行

### 8.2 覆盖率目标

| 模块 | 目标覆盖率 |
|------|-----------|
| DTO 层 | ≥ 90% |
| Service 层 | ≥ 80% |
| Namespace 层 | ≥ 70% |
| 整体 | ≥ 75% |

### 8.3 CI/CD 验收

- [ ] GitHub Actions 能够成功运行所有测试
- [ ] 测试失败时 PR 无法合并
- [ ] 测试执行时间 < 5 分钟

---

## 附录

### A. 参考资料

- [Mimesis Documentation](https://mimesis.name/)
- [Polyfactory Documentation](https://polyfactory.litestar.dev/)
- [python-socketio Testing](https://python-socketio.readthedocs.io/en/latest/client.html)
- [Contract Testing](https://martinfowler.com/bliki/ContractTest.html)

### B. 相关文档

- `docs/socketio_event_development_standard.md` - Socket.IO 事件开发规范
- `manual_tests/MANUAL_TEST.md` - 手动测试指南
- `.github/workflows/tests.yml` - CI/CD 配置

---

**维护者**: JQQ <jqq1716@gmail.com>  
**最后更新**: 2026-01-08
