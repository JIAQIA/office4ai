---
description: 更新现有的 Socket.IO 事件
argument-hint: <event_name> 例如 word:get:documentContent
---

# Socket.IO 事件更新 Skill

根据事件名 `$1` 更新现有的 Socket.IO 事件实现。更新事件比创建事件更复杂，需要：

1. **协议变更分析** - 仔细对比新旧协议差异
2. **代码实现更新** - 修改 DTO、Namespace、Service
3. **测试更新与验证** - 新功能验证、旧功能保持、废弃点清理

## 第一阶段：协议变更分析

### 1.1 获取协议文档

用户通常会提供 Confluence 文档链接，使用 MCP 工具阅读并提取：

- **新增字段**：哪些参数/响应字段是新增的
- **修改字段**：哪些字段的类型、约束、默认值发生变化
- **废弃字段**：哪些字段被标记为废弃或删除
- **行为变更**：事件的处理逻辑是否有变化

### 1.2 定位现有实现

```
DTO 定义      → office4ai/environment/workspace/dtos/<platform>.py
Namespace     → office4ai/environment/workspace/socketio/namespaces/<platform>.py
单元测试      → tests/unit_tests/office4ai/environment/workspace/dtos/
契约测试      → tests/contract_tests/word/
手动测试      → manual_tests/<event_name>_e2e/
```

### 1.3 生成变更清单

创建变更对照表（示例）：

| 变更类型 | 字段/行为 | 旧值 | 新值 | 影响范围 |
|---------|----------|-----|-----|---------|
| 新增 | `options.maxLength` | - | `int \| None` | DTO, 测试 |
| 修改 | `format.color` | `str` | `str \| None` | DTO, 测试 |
| 废弃 | `legacyParam` | 存在 | 移除 | DTO, 测试 |
| 行为 | 空选区处理 | 返回错误 | 返回空数据 | Namespace, 测试 |

## 第二阶段：代码实现更新

### 2.1 更新 DTO

**新增字段**：
```python
# 在现有 Request 类中添加
new_field: Optional[str] = Field(None, alias="newField", description="新字段描述")
```

**修改字段**：
```python
# 修改类型或约束
old_field: str | None = Field(None, alias="oldField")  # 原来是 str
```

**废弃字段**：
```python
# 方式1: 添加 deprecated 标记（保持向后兼容）
legacy_field: Optional[str] = Field(
    None, 
    alias="legacyField", 
    deprecated=True,
    description="[DEPRECATED] 请使用 newField"
)

# 方式2: 直接删除（破坏性变更）
# 删除字段定义
```

### 2.2 更新 Namespace（如有行为变更）

```python
async def on_<event_name>(self, sid: str, data: Any) -> None:
    # 更新处理逻辑
    # 注意保持向后兼容性
```

### 2.3 更新响应 DTO（如有）

同样遵循新增/修改/废弃的处理方式。

## 第三阶段：测试更新与验证

### 3.1 测试更新策略

| 变更类型 | 测试策略 |
|---------|---------|
| **新功能** | 添加新测试用例覆盖新字段/行为 |
| **旧功能** | 确保现有测试仍然通过（回归保护） |
| **废弃点** | 移除相关测试或标记为跳过 |

### 3.2 单元测试更新

位置：`tests/unit_tests/office4ai/environment/workspace/dtos/test_<platform>_dtos.py`

```python
class Test<Platform><Action>Request:
    # 新增测试
    def test_new_field_valid(self):
        request = Request(requestId="req-1", documentUri="...", newField="value")
        assert request.new_field == "value"
    
    def test_new_field_optional(self):
        request = Request(requestId="req-1", documentUri="...")
        assert request.new_field is None
    
    # 保持现有测试（确保回归）
    def test_existing_behavior(self):
        # ... 现有测试保持不变
```

### 3.3 契约测试更新

位置：`tests/contract_tests/word/test_<event_name>.py`

```python
@pytest.mark.asyncio
@pytest.mark.contract
async def test_new_feature(workspace, mock_addin_client):
    # 注册包含新字段的响应
    mock_addin_client.register_response("event_name", lambda req: {
        "requestId": req["requestId"],
        "newField": "expected_value"
    })
    # ... 执行并验证
```

### 3.4 手动测试更新

位置：`manual_tests/<event_name>_e2e/`

**更新 README.md**：
- 添加新测试场景到测试矩阵
- 标记废弃的测试场景
- 更新最后修改日期

**更新测试文件**：
- 新增测试函数覆盖新参数组合
- 移除废弃参数的测试
- 更新 `TEST_MAPPING` 字典

详细的手动测试更新指南见：[reference.md](./reference.md)

## 验证检查清单

### 代码验证
```bash
poe lint          # 代码检查
poe typecheck     # 类型检查
```

### 测试验证
```bash
poe test-unit     # 单元测试 - 验证 DTO 变更
poe test-contract # 契约测试 - 验证协议兼容性
```

### 手动验证
```bash
# 运行相关手动测试
uv run python manual_tests/<event_name>_e2e/test_xxx.py --test all
```

## 向后兼容性指南

| 场景 | 处理方式 |
|-----|---------|
| 新增可选字段 | ✅ 安全，直接添加 |
| 新增必填字段 | ⚠️ 破坏性，需要版本协调 |
| 修改字段类型 | ⚠️ 可能破坏性，评估影响 |
| 删除字段 | ⚠️ 破坏性，先标记废弃 |
| 修改默认值 | ⚠️ 可能影响行为，需测试 |

## 关键文件引用

- **DTO 基类**: `office4ai/environment/workspace/dtos/common.py`
- **错误码**: `office4ai/environment/workspace/dtos/common.py` → `ErrorCode`
- **测试 Fixtures**: `tests/contract_tests/conftest.py`
- **Mock Client**: `tests/contract_tests/mock_addin/client.py`

## 完成标准

- [ ] 协议变更已完整分析并记录
- [ ] DTO 已更新（新增/修改/废弃）
- [ ] Namespace 已更新（如有行为变更）
- [ ] 单元测试已更新并通过
- [ ] 契约测试已更新并通过
- [ ] 手动测试 README 已更新
- [ ] 手动测试用例已更新
- [ ] 所有验证命令通过
