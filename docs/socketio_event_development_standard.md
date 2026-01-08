# Socket.IO 事件开发规范

> **版本**: v1.0.0
> **最后更新**: 2026-01-08
> **状态**: 正式生效
> **适用范围**: Office4AI Workspace Socket.IO 事件开发

---

## 1. 总体原则

### 1.1 核心理念

- **文档优先**: 所有开发以 Confluence 协议文档为第一标准
- **分层架构**: 严格遵循 Namespace → Service → 业务逻辑的分层设计
- **类型安全**: 所有数据传输必须通过强类型 DTO
- **测试驱动**: 新事件必须包含完整的单元测试和按需的集成测试

### 1.2 团队协作

- **大型团队规范**: 适用于 15+ 人团队，强调一致性和可维护性
- **代码审查**: 所有事件实现必须经过 Code Review
- **文档同步**: 代码变更与协议文档保持同步

---

## 2. 架构分层

### 2.1 分层设计

```
┌─────────────────────────────────────┐
│   Socket.IO Protocol Layer          │  Confluence 文档定义
│   (word:insert:text, etc.)           │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│   Namespace Layer                    │  事件路由、连接管理
│   - WordNamespace                    │  office4ai/environment/workspace/
│   - PptNamespace                     │    socketio/namespaces/
│   - ExcelNamespace                   │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│   DTO Layer (Data Transfer Objects)  │  数据模型、类型定义
│   - BaseRequest/BaseResponse         │  office4ai/environment/workspace/
│   - Event-specific DTOs              │    dtos/
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│   Service Layer                      │  业务逻辑实现
│   - WordService                      │  office4ai/environment/workspace/
│   - PptService                       │    socketio/services/
│   - ExcelService                     │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│   Office API Layer                   │  LibreOffice UNO Bridge
│   (future implementation)            │
└─────────────────────────────────────┘
```

### 2.2 职责划分

| 层级 | 职责 | 不涉及 |
|------|------|--------|
| **Namespace** | 事件路由、连接管理、请求包装、异常捕获 | 业务逻辑、Office API 调用 |
| **DTO** | 数据结构定义、类型验证、格式转换 | 业务逻辑 |
| **Service** | 业务逻辑实现、Office API 封装 | Socket.IO 通信细节 |
| **Office API** | 文档操作、内容处理 | 网络通信 |

---

## 3. 开发工作流

### 3.1 标准流程

```
1. 阅读协议文档 (Confluence)
   ↓
2. 定义 DTO (dtos/*.py)
   ↓
3. 实现 Service (services/*.py)
   ↓
4. 实现 Namespace 事件处理器 (namespaces/*.py)
   ↓
5. 编写单元测试 (tests/unit_tests/*)
   ↓
6. 编写集成测试 (tests/integration_tests/*)
   ↓
7. Code Review
   ↓
8. 更新协议文档（如有变更）
```

### 3.2 文档变更流程

当发现协议文档定义不合理时：

1. **创建 Issue**: 在项目 Issue Tracker 中提出问题
2. **团队评审**: 与团队成员讨论修改方案
3. **文档优先**: 先更新 Confluence 协议文档
4. **代码实现**: 按新文档实现代码
5. **同步确认**: 确保 Confluence 文档与代码一致

**禁止**未经评审直接修改协议文档或按个人理解偏离文档实现。

---

## 4. DTO 定义规范

### 4.1 文件组织

```
office4ai/environment/workspace/dtos/
├── common.py              # 通用类型（BaseRequest, BaseResponse, ErrorResponse）
├── word.py                # Word 事件 DTO
├── ppt.py                 # PPT 事件 DTO
└── excel.py               # Excel 事件 DTO
```

### 4.2 命名约定

#### 事件名称
- **格式**: `<platform>:<action>[:<resource>]`
- **示例**:
  - `word:get:selectedContent`
  - `ppt:insert:text`
  - `excel:set:cellValue`

#### DTO 类名
- **请求 DTO**: `<Platform><Action>Request`
  - 示例: `WordGetSelectedContentRequest`, `PptInsertTextRequest`
- **嵌套选项**: `<Action><Resource>Options`
  - 示例: `GetContentOptions`, `InsertTextOptions`
- **响应数据**: `<Resource>Result` 或 `<Resource>Data`
  - 示例: `SelectedContentResult`, `DocumentStructureData`

#### 字段命名
- **内部 (Python)**: `snake_case` (document_uri, max_text_length)
- **外部 (Socket.IO)**: `camelCase` (documentUri, maxTextLength)
- **映射方式**: 使用 Pydantic `alias` 参数

### 4.3 定义规范

#### 基础结构

所有请求 DTO 必须继承 `BaseRequest`：

```python
from office4ai.environment.workspace.socketio.dtos.common import BaseRequest

class WordGetSelectedContentRequest(BaseRequest):
    """获取选中内容请求

    协议文档: https://turingfocus.atlassian.net/...
    """

    # 必须定义 event_name（自动注册）
    event_name: ClassVar[str] = "word:get:selectedContent"

    # 业务字段定义
```

#### 字段定义要求

1. **类型注解**: 所有字段必须有明确类型
2. **Alias 映射**: 外部字段名使用 `alias`
3. **文档字符串**: 复杂字段必须有 description
4. **验证规则**: 使用 Pydantic 验证（min_length, ge, le 等）
5. **可选标记**: 使用 `Optional[T]` 或 `T | None`

#### 嵌套类型组织

- **简单类型**: 直接在同一文件中定义
- **复用类型**: 定义在通用类型文件中（如 `common.py`）
- **复杂嵌套**: 评估复用性后决定独立文件或同文件

### 4.4 自动注册机制

所有继承 `BaseRequest` 的 DTO 会自动注册到全局注册表：

```python
# 自动注册，无需手动操作
class WordMyEventRequest(BaseRequest):
    event_name: ClassVar[str] = "word:my:event"
    # ...
```

注册后可通过 `RequestRegistry` 访问。

---

## 5. Service 层实现规范

### 5.1 文件组织

```
office4ai/environment/workspace/socketio/services/
├── word_service.py        # WordService
├── ppt_service.py         # PptService
└── excel_service.py       # ExcelService
```

**注意**: BaseService 抽象基类待实现。

### 5.2 类组织

按应用分组 Service：
- `WordService`: 所有 Word 事件业务逻辑
- `PptService`: 所有 PPT 事件业务逻辑
- `ExcelService`: 所有 Excel 事件业务逻辑

### 5.3 方法签名

**要求**: 使用业务参数，而非 DTO

```python
class WordService:
    async def get_selected_content(
        self,
        document_uri: str,
        include_text: bool = True,
        include_images: bool = True,
        include_tables: bool = True,
        max_text_length: int | None = None,
        timeout: int | None = None,
    ) -> dict[str, Any]:
        """获取选中内容

        Args:
            document_uri: 文档 URI
            include_text: 是否包含文本
            include_images: 是否包含图片
            include_tables: 是否包含表格
            max_text_length: 最大文本长度限制
            timeout: 超时时间（秒），None 使用默认值

        Returns:
            业务结果字典

        Raises:
            OfficeApiError: Office API 调用失败
            ValidationError: 参数验证失败
        """
```

**设计原则**:
- 方法名与事件动作对应（get_selected_content）
- 参数拆解为业务字段，而非整个 Request DTO
- 返回业务结果字典，不包含 request_id、success 等响应包装字段
- 异常向上传播，由 Namespace 层统一处理
- timeout 参数可选，支持动态超时配置

### 5.4 错误处理

Service 层抛出业务异常，不捕获：

```python
# 定义业务异常
class OfficeApiError(Exception):
    """Office API 调用失败"""
    def __init__(self, code: str, message: str, details: dict = None):
        self.code = code
        self.message = message
        self.details = details
        super().__init__(message)

# Service 方法中抛出异常
async def insert_text(self, document_uri: str, text: str, ...):
    if not text:
        raise ValidationError("text cannot be empty")

    try:
        result = await office_api.insert_text(...)
    except OfficeException as e:
        raise OfficeApiError("3000", "Failed to insert text", {"error": str(e)})
```

---

## 6. Namespace 事件处理器规范

### 6.1 文件组织

```
office4ai/environment/workspace/socketio/namespaces/
├── base.py                # BaseNamespace 抽象基类
├── word.py                # WordNamespace
├── ppt.py                 # PptNamespace
└── excel.py               # ExcelNamespace
```

### 6.2 事件处理器注册

**按命名约定自动注册**:

```python
class WordNamespace(BaseNamespace):
    async def on_word_get_selectedContent(
        self,
        sid: str,
        data: dict[str, Any],
    ):
        """处理 word:get:selectedContent 事件

        事件名: word:get:selectedContent
        协议文档: https://turingfocus.atlassian.net/...
        """
        # 实现
```

**命名规则**: `on_<event_name_with_underscores>`

- `word:get:selectedContent` → `on_word_get_selectedContent`
- `ppt:insert:text` → `on_ppt_insert_text`
- `excel:set:cellValue` → `on_excel_set_cell_value`

### 6.3 事件处理器实现

#### 标准结构

```python
async def on_word_get_selectedContent(self, sid: str, data: dict[str, Any]):
    """处理 word:get:selectedContent 事件"""
    # 1. 解析请求
    request = WordGetSelectedContentRequest(**data)

    # 2. 调用 Service
    try:
        result = await self.word_service.get_selected_content(
            document_uri=request.document_uri,
            include_text=request.options.include_text if request.options else True,
            timeout=getattr(request, 'timeout', None),
        )
        # 3. 返回成功响应
        await self.emit_success_response(sid, request.request_id, result)

    except ValidationError as e:
        # 4. 处理验证错误 (4xxx)
        await self.emit_error_response(
            sid, request.request_id, "4000", str(e)
        )
    except OfficeApiError as e:
        # 5. 处理 Office API 错误 (3xxx)
        await self.emit_error_response(
            sid, request.request_id, e.code, e.message, e.details
        )
    except Exception as e:
        # 6. 处理未知错误 (1xxx)
        logger.exception(f"Unexpected error in word:get:selectedContent: {e}")
        await self.emit_error_response(
            sid, request.request_id, "1000", "Unknown error"
        )
```

#### 职责边界

- **负责**:
  - 解析请求 DTO
  - 调用 Service 方法
  - 捕获异常并转换为错误响应
  - 发送响应到客户端

- **不负责**:
  - 业务逻辑
  - Office API 调用
  - 复杂数据处理

### 6.4 响应发送

使用 `BaseNamespace` 提供的辅助方法：

```python
# 成功响应
await self.emit_success_response(
    sid,
    request_id="req-123",
    data={"text": "...", "elements": [...]},
)

# 错误响应
await self.emit_error_response(
    sid,
    request_id="req-123",
    code="3002",
    message="Selection is empty",
    details={"hint": "Please select some text first"},
)
```

### 6.5 客户端上报事件 (C→S)

对于 fire-and-forget 事件：

```python
async def on_word_event_selectionChanged(self, sid: str, data: dict[str, Any]):
    """处理 word:event:selectionChanged 事件"""
    # 1. 记录事件
    client_info = self.connection_manager.get_client(sid)
    logger.info(
        f"Selection changed: client={client_info.client_id}, "
        f"document={client_info.document_uri}, data={data}"
    )

    # 2. 更新状态（如需要）

    # 3. 不发送响应（fire-and-forget）
    # 除非文档明确标记为"需要确认"的事件
```

---

## 7. 错误处理规范

### 7.1 错误码体系

| 类别 | 范围 | 说明 |
|------|------|------|
| 通用错误 | 1xxx | 系统级错误 |
| 认证错误 | 2xxx | 认证和授权 |
| Office API 错误 | 3xxx | AddIn 端返回的错误 |
| 验证错误 | 4xxx | 请求参数验证失败 |

**参考文档**: 错误码参考 (Confluence)

### 7.2 错误响应格式

AddIn 端错误通过 `success=false` + `error` 对象返回：

```python
{
    "requestId": "req-123",
    "success": false,
    "error": {
        "code": "3002",
        "message": "Selection is empty",
        "details": {"hint": "Please select some text first"}
    },
    "timestamp": 1704700800000
}
```

### 7.3 异常分类处理

```python
try:
    result = await service.method(...)
except ValidationError as e:
    # 参数验证失败 → 4xxx
    await self.emit_error_response(sid, request_id, "4000", str(e))
except OfficeApiError as e:
    # Office API 错误 → 使用错误自带码 (3xxx)
    await self.emit_error_response(sid, request_id, e.code, e.message, e.details)
except TimeoutError as e:
    # 超时 → 1002
    await self.emit_error_response(sid, request_id, "1002", "Request timeout")
except Exception as e:
    # 未知错误 → 1000
    logger.exception(f"Unexpected error: {e}")
    await self.emit_error_response(sid, request_id, "1000", "Unknown error")
```

### 7.4 超时配置

支持动态超时参数：

```python
# Service 方法
async def method(self, ..., timeout: int | None = None):
    # timeout: None → 使用默认超时（如 10 秒）
    # timeout: 30 → 使用 30 秒超时
```

---

## 8. 测试规范

### 8.1 测试要求

| 事件类型 | 单元测试 | 集成测试 | 说明 |
|----------|----------|----------|------|
| 简单查询类 | ✅ 必需 | ❌ 可选 | DTO 验证、Service 方法 |
| 复杂操作类 | ✅ 必需 | ✅ 必需 | 包含完整事件流程 |
| 客户端上报事件 | ✅ 必需 | ❌ 可选 | 事件接收、状态更新 |

### 8.2 单元测试

#### DTO 测试

```python
class TestWordGetSelectedContentRequest:
    def test_valid_request(self):
        """测试创建有效请求"""
        request = WordGetSelectedContentRequest(
            requestId="req-123",
            documentUri="file:///test.docx",
            options={"includeText": True}
        )
        assert request.request_id == "req-123"

    def test_validation_error(self):
        """测试验证错误"""
        with pytest.raises(ValidationError):
            WordGetSelectedContentRequest(
                requestId="req-123",
                # 缺少必填字段 documentUri
            )
```

#### Service 测试

```python
class TestWordService:
    async def test_get_selected_content_success(self):
        """测试成功获取选中内容"""
        result = await word_service.get_selected_content(
            document_uri="file:///test.docx",
            include_text=True,
        )
        assert "text" in result
        assert "elements" in result

    async def test_get_selected_content_empty_selection(self):
        """测试选择为空的情况"""
        with pytest.raises(OfficeApiError) as exc_info:
            await word_service.get_selected_content(
                document_uri="file:///test.docx",
            )
        assert exc_info.value.code == "3002"
```

### 8.3 集成测试

#### 测试 Fixture

提供统一的 `mock_addin_client` fixture：

```python
@pytest_asyncio.fixture
async def mock_addin_client(socketio_server):
    """模拟 AddIn 客户端，提供预定义的响应"""
    client = await socketio_client.connect()

    # Mock 响应
    async def mock_handler(request, callback):
        if request.get("options", {}).get("includeText"):
            callback({
                "requestId": request["requestId"],
                "success": True,
                "data": {"text": "mocked content", "elements": []},
                "timestamp": int(time.time() * 1000),
            })
        else:
            callback({
                "requestId": request["requestId"],
                "success": False,
                "error": {"code": "4000", "message": "Invalid options"},
                "timestamp": int(time.time() * 1000),
            })

    client.on('word:get:selectedContent', mock_handler)
    yield client
    await client.disconnect()
```

#### 集成测试用例

```python
async def test_get_selected_content_integration(mock_addin_client):
    """测试完整的事件流程"""
    # 发送请求
    response = await mock_addin_client.emitWithAck(
        'word:get:selectedContent',
        {
            'requestId': 'req-123',
            'documentUri': 'file:///test.docx',
            'options': {'includeText': True}
        }
    )

    # 验证响应
    assert response['success'] is True
    assert 'data' in response
    assert response['data']['text'] == 'mocked content'
```

### 8.4 测试组织

```
tests/
├── unit_tests/
│   └── office4ai/environment/workspace/socketio/
│       ├── dtos/              # DTO 测试
│       ├── services/          # Service 测试
│       └── namespaces/        # Namespace 测试
└── integration_tests/
    └── office4ai/environment/workspace/socketio/
        ├── test_word_events.py    # Word 事件集成测试
        ├── test_ppt_events.py     # PPT 事件集成测试
        └── test_excel_events.py   # Excel 事件集成测试
```

---

## 9. 文档管理规范

### 9.1 协议文档结构

**Confluence 空间**: OFFICE4AI-Workspace-Socket.IO

```
OFFICE4AI-Workspace-Socket.IO
├── Socket.IO API 标准 (首页)
├── 01-基础规范/
│   ├── 事件命名规范
│   ├── 数据结构定义
│   ├── 错误码参考
│   └── 错误处理最佳实践
├── 02-Word 事件/
│   ├── Word 事件索引
│   ├── word:get:selectedContent
│   ├── word:insert:text
│   └── ...
├── 03-PowerPoint 事件/
├── 04-Excel 事件/
├── 05-客户端上报事件/
└── 变更日志
```

### 9.2 事件文档必需内容

每个事件文档必须包含：

1. **基本信息**: 状态、方向、命名空间
2. **概述**: 功能说明
3. **请求结构**: TypeScript + Python 类型定义
4. **响应结构**: TypeScript + Python 类型定义
5. **可能错误**: 相关错误码列表
6. **注意事项**: 重要提示
7. **使用示例**: 服务器端和客户端示例代码
8. **标签**: 便于搜索和分类
9. **超时要求**: 是否需要 ACK 确认（仅客户端上报事件）

### 9.3 文档一致性检查

- **Code Review**: 检查代码实现与协议文档一致性
- **手动检查**: 对比 DTO 定义与文档中的类型定义
- **变更同步**: 协议文档变更后，相关代码必须同步更新

### 9.4 版本管理

- **向后兼容**: 新版本不破坏旧事件
- **迁移机制**: 废弃事件标记 `deprecated`，保留至少一个版本周期
- **变更日志**: 记录所有协议变更到变更日志页面

---

## 10. 代码组织规范

### 10.1 导入顺序

```python
# 1. 标准库
import asyncio
from typing import Any

# 2. 第三方库
from pydantic import Field
import pytest

# 3. 本地模块
from office4ai.environment.workspace.dtos.common import BaseRequest
from office4ai.environment.workspace.socketio.services.word_service import WordService
```

### 10.2 文件命名

- **模块文件**: `snake_case.py` (word_service.py, base.py)
- **测试文件**: `test_<module>.py` (test_word_service.py)
- **DTO 文件**: 按应用命名 (word.py, ppt.py, excel.py)
- **命名空间文件**: 按应用命名 (word.py, ppt.py, excel.py)

### 10.3 代码风格

- **行长度**: 120 字符
- **类型注解**: 强制要求 (mypy `disallow_untyped_defs = true`)
- **文档字符串**: 公共类和方法必须有 docstring
- **命名规范**:
  - 类: `PascalCase`
  - 方法/函数: `snake_case`
  - 常量: `UPPER_SNAKE_CASE`

---

## 11. 最佳实践

### 11.1 开发原则

1. **文档第一**: 实现前先阅读协议文档
2. **类型安全**: 充分利用 Pydantic 类型验证
3. **异常传播**: Service 层抛异常，Namespace 层统一处理
4. **测试覆盖**: 核心逻辑必须有单元测试
5. **日志记录**: 关键操作和异常必须记录日志

### 11.2 性能考虑

1. **异步优先**: 所有 I/O 操作使用 async/await
2. **超时控制**: 长时间操作必须设置超时
3. **连接复用**: Service 实例复用，避免重复创建

### 11.3 安全考虑

1. **输入验证**: 所有客户端输入必须验证
2. **错误信息**: 不暴露敏感系统信息
3. **URI 验证**: document_uri 必须验证合法性

### 11.4 可维护性

1. **代码注释**: 复杂逻辑添加注释说明
2. **文档更新**: 功能变更及时更新协议文档
3. **重构原则**: 保持简单，避免过度设计

---

## 12. 常见问题

### Q1: 如何处理协议文档中没有定义的情况？

**A**: 创建 Issue 讨论，由团队评审后更新协议文档，再按新文档实现。禁止私自偏离文档。

### Q2: Service 层方法应该返回什么类型？

**A**: 返回业务结果字典（`dict[str, Any]`），不包含 request_id、success 等响应包装字段。

### Q3: 如何测试需要 AddIn 配合的事件？

**A**: 使用集成测试的 `mock_addin_client` fixture，模拟 AddIn 响应。

### Q4: 客户端上报事件需要发送响应吗？

**A**: 一般不需要（fire-and-forget）。除非协议文档明确标记为"需要确认"。

### Q5: 如何添加新的错误码？

**A**:
1. 在 Issue 中提出新错误码需求
2. 评审通过后更新 Confluence"错误码参考"页面
3. 在代码中使用新错误码

---

## 附录

### A. 参考文档

- **Socket.IO API 标准**: https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI/overview
- **项目 CLAUDE.md**: /Users/jqq/PycharmProjects/office4ai/CLAUDE.md
- **现有实现**: `office4ai/environment/workspace/socketio/namespaces/word.py`

### B. 工具链

- **代码检查**: `poe lint`, `poe typecheck`
- **测试运行**: `poe test`, `poe test-unit`, `poe test-integration`
- **格式化**: `poe format`

### C. 代码示例参考

- **已实现事件**: `word:get:selectedContent`
- **DTO 定义**: `office4ai/environment/workspace/socketio/dtos/word.py`
- **测试用例**: `tests/unit_tests/office4ai/environment/workspace/socketio/dtos/test_word.py`

---

**维护者**: JQQ <jqq1716@gmail.com>
**最后更新**: 2026-01-08
**下次评审**: 2026-02-01
