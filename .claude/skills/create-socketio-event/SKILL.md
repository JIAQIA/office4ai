---
description: 开发新的 Socket.IO 事件。根据事件名和 OASP 协议规范，生成完整的 DTO、Namespace 处理器和单元测试。
argument-hint: <event_name> 例如 word:get:documentContent
---

# Socket.IO 事件开发 Skill

根据事件名实现完整的 Socket.IO 事件，遵循项目既有模式。

## 核心思路

### 为什么需要此 Skill

事件开发涉及三个分层：DTO（数据结构）、Namespace（事件路由）、测试。每层都有既定模式，需要保持一致性：

- **DTO 层**：定义请求/响应结构，自动注册到 `request_registry`
- **Namespace 层**：处理事件，遵循命名规则 `on_<event_with_underscores>`
- **测试层**：验证 DTO 序列化、验证规则、边界情况

### 架构分层

```
DTO Layer (数据结构)     → office4ai/environment/workspace/dtos/<platform>.py
Namespace Layer (事件路由) → office4ai/environment/workspace/socketio/namespaces/<platform>.py
Test Layer (单元测试)     → tests/unit_tests/office4ai/environment/workspace/dtos/
```

## 执行步骤

### Step 1: 查找事件协议定义

根据事件名（如 `word:get:documentContent`）查找 OASP 协议规范。

**查找顺序**：

1. **本地 oasp-protocol 目录**（推荐）：如果项目 workspace 中管理了 `oasp-protocol` 目录，直接阅读对应的 `.md` 文件
2. **在线文档**：访问 https://a2c-smcp.github.io/oasp-protocol/latest/ 获取定义
   - 如用户指定版本（如 `v1.0`），使用对应版本 URL

**从规范中提取**：
- 事件名称格式：`<platform>:<action>:<target>`
- 请求参数（除 `requestId`、`documentUri`、`timestamp` 外的业务字段）
- 响应数据结构
- 可能的错误码

### Step 2: 定义 DTO

在对应平台的 DTO 文件中添加请求类。

**参考模式**：[office4ai/environment/workspace/dtos/word.py](office4ai/environment/workspace/dtos/word.py)

**关键模式**：

```python
# 1. 继承 BaseRequest 自动注册
class <Platform><Action>Request(BaseRequest):
    event_name: ClassVar[str] = "<platform>:<action>:<target>"

    # 2. 字段用 snake_case，alias 映射 camelCase
    field_name: str = Field(..., alias="fieldName", description="...")
```

**DTO 规范**：
- 必须定义 `event_name: ClassVar[str]` 触发自动注册
- 嵌套类型继承 `SocketIOBaseModel`
- 可选字段设置 `default=None`
- 参考 [common.py](office4ai/environment/workspace/dtos/common.py) 中的 `BaseRequest` 和 `SocketIOBaseModel`

### Step 3: 实现 Namespace 处理器

在对应平台的 Namespace 文件中添加事件处理器。

**参考模式**：[office4ai/environment/workspace/socketio/namespaces/word.py](office4ai/environment/workspace/socketio/namespaces/word.py)

**命名规则**：`on_<event_name>` 其中 `:` 替换为 `_`

```python
async def on_word_get_documentContent(self, sid: str, data: Any) -> None:
    """
    处理 word:get:documentContent 事件

    Event: word:get:documentContent
    Direction: Client → Server
    """
    client_info = self.get_client_info(sid)
    if client_info:
        logger.info(f"Received word:get:documentContent from {client_info.client_id}")
```

### Step 4: 编写单元测试

为 DTO 添加完整的单元测试。

**参考模式**：[tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py](tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py)

**测试覆盖**：
- `test_valid_request_with_defaults` - 默认值创建
- `test_missing_required_fields` - 必填字段验证
- `test_event_name_attribute` - 事件名类变量
- `test_to_payload_camel_case` - 序列化别名
- `test_build_class_method` - build 工厂方法

### Step 5: 添加手动测试（可选但推荐）

对于需要真实 Office 环境验证的事件，添加 Manual Tests。

**参考模式**：[manual_tests/insert_text_e2e/](manual_tests/insert_text_e2e/) 目录结构

**目录结构**：
```
manual_tests/<event_name>_e2e/
├── README.md              # ⭐ 必需：测试说明文档
├── test_basic_xxx.py      # 基础功能测试
├── test_<dimension>.py    # 按维度分组（如 location、format）
└── test_edge_cases.py     # 边界情况测试
```

**何时需要 Manual Tests**：
- 涉及 Office UI 交互（如选区、光标位置）
- 需要验证格式参数的视觉效果
- 边界情况难以自动化模拟

**详细指南**：参考 [test-checklist Skill](.claude/skills/test-checklist/SKILL.md) 中的 Manual Tests 章节和 [reference.md](.claude/skills/test-checklist/reference.md) 中的参数组合矩阵。

### Step 6: 验证

```bash
poe lint          # 代码检查
poe typecheck     # 类型检查
poe test-unit     # 单元测试
```

## 错误码参考

| 范围 | 类别 | 定义位置 |
|------|------|---------|
| 1xxx | 通用错误 | [ErrorCode](office4ai/environment/workspace/dtos/common.py) |
| 2xxx | 认证错误 | 同上 |
| 3xxx | Office API 错误 | 同上 |
| 4xxx | 验证错误 | 同上 |

## 关键类参考

- [BaseRequest](office4ai/environment/workspace/dtos/common.py) - 请求基类
- [SocketIOBaseModel](office4ai/environment/workspace/dtos/common.py) - 嵌套类型基类
- [BaseNamespace](office4ai/environment/workspace/socketio/namespaces/base.py) - Namespace 基类
- [ErrorCode](office4ai/environment/workspace/dtos/common.py) - 错误码常量
