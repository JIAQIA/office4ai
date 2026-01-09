---
description: 开发新的 Socket.IO 事件
argument-hint: <event_name> 例如 word:get:documentContent
---

# Socket.IO 事件开发 Skill

根据事件名 `$1` 实现完整的 Socket.IO 事件。一般而言用户会提供一个Confluence文档，你借助MCP工具来阅读文档并获取以下信息：

1. **事件名称**（格式：`<platform>:<action>:<target>`）
   - 平台：`word`, `ppt`, `excel`, `common`
   - 操作：`get`, `insert`, `replace`, `update`, `delete`, `event`
   - 目标：如 `selectedContent`, `text`, `table`

2. **功能描述**：一句话说明这个事件的作用

3. **请求参数**：需要哪些额外参数（除 `requestId` 和 `documentUri` 外）

4. **响应数据**：返回什么数据

5. **可能的错误**：会抛出哪些特定错误

如果用户未提供有效文档或者发生异常，你可以借助 AskUserQuestionTool 来获取相应内容

## 架构分层

```
Namespace Layer (事件路由)     → office4ai/environment/workspace/socketio/namespaces/
DTO Layer (数据结构)           → office4ai/environment/workspace/dtos/
Service Layer (业务逻辑)       → office4ai/environment/workspace/socketio/services/
```

## 开发步骤

### 1. 定义 DTO

在 `office4ai/environment/workspace/dtos/<platform>.py` 中定义请求 DTO：

```python
from typing import ClassVar, Optional
from pydantic import Field
from .common import BaseRequest, SocketIOBaseModel

class <Platform><Action>Request(BaseRequest):
    """<事件描述>"""
    
    event_name: ClassVar[str] = "<platform>:<action>:<resource>"
    
    # 业务字段 (使用 alias 映射 camelCase)
    field_name: str = Field(..., alias="fieldName", description="字段描述")
```

**DTO 规范**:
- 继承 `BaseRequest`（自动注册到 `request_registry`）
- 必须定义 `event_name: ClassVar[str]`
- 字段使用 `snake_case`，通过 `alias` 映射到 `camelCase`
- 嵌套类型继承 `SocketIOBaseModel`

### 2. 实现 Namespace 事件处理器

在 `office4ai/environment/workspace/socketio/namespaces/<platform>.py` 中添加处理器：

```python
async def on_<event_name_with_underscores>(self, sid: str, data: Any) -> None:
    """
    处理 <event_name> 事件
    
    Event: <platform>:<action>:<resource>
    Direction: Client → Server
    """
    client_info = self.get_client_info(sid)
    if client_info:
        logger.info(f"Received <event_name> from {client_info.client_id}")
```

**命名规则**: `on_<event_name>` 其中 `:` 替换为 `_`
- `word:get:selectedContent` → `on_word_get_selectedContent`
- `ppt:insert:text` → `on_ppt_insert_text`

### 3. 编写单元测试

在 `tests/unit_tests/office4ai/environment/workspace/socketio/dtos/` 中添加测试：

```python
class Test<Platform><Action>Request:
    def test_valid_request(self):
        request = <Platform><Action>Request(
            requestId="req-123",
            documentUri="file:///test.docx",
            # 业务字段...
        )
        assert request.request_id == "req-123"
    
    def test_validation_error(self):
        with pytest.raises(ValidationError):
            <Platform><Action>Request(requestId="req-123")  # 缺少必填字段
```

## 错误码体系

| 范围 | 类别 | 示例 |
|------|------|------|
| 1xxx | 通用错误 | 1000=未知错误, 1002=超时 |
| 2xxx | 认证错误 | 2000=未授权 |
| 3xxx | Office API 错误 | 3002=选区为空 |
| 4xxx | 验证错误 | 4000=参数验证失败 |

## 关键类引用

- `BaseRequest`: `office4ai/environment/workspace/dtos/common.py`
- `SocketIOBaseModel`: `office4ai/environment/workspace/dtos/common.py`
- `BaseNamespace`: `office4ai/environment/workspace/socketio/namespaces/base.py`
- `ErrorCode`: `office4ai/environment/workspace/dtos/common.py`

## 验证命令

```bash
poe lint          # 代码检查
poe typecheck     # 类型检查
poe test-unit     # 单元测试
```
