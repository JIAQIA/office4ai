---
name: testing-checklist
description: 开发新功能时的测试清单，确保不遗漏任何测试层级
usage: |-
  开发新功能时参考此清单，确保完成所有必要测试。
  如需详细代码示例，请参阅 TESTING_STANDARD.md
---
# Test Architecture Skill

## 测试分层架构

```
Unit Tests (tests/unit_tests/)     → DTO验证、Service逻辑 [CI ✅]
Integration Tests (tests/integration_tests/) → Socket.IO服务器+客户端 [CI ✅]
Contract Tests (tests/contract_tests/)  → Mock Add-In完整流程 [CI ✅]
Manual Tests (manual_tests/)       → 真实Office E2E验证 [CI ❌]
```

## MockAddInClient 核心接口

```python
class MockAddInClient:
    async def connect(self) -> None: ...
    async def disconnect(self) -> None: ...
    def register_response(self, event: str, factory: Callable[[dict], dict]) -> None: ...
    def register_static_response(self, event: str, response: dict) -> None: ...
    @property
    def received_events(self) -> list[tuple[str, dict]]: ...
```

## 响应注册模式

```python
# 静态响应
mock_client.register_static_response("word:get:selectedContent", {"success": True, "data": {...}})

# 动态响应
mock_client.register_response("word:get:selectedContent", lambda req: {"requestId": req["requestId"], ...})
```

## Contract Test 模板

```python
@pytest.mark.asyncio
@pytest.mark.contract
async def test_event_success(workspace, mock_addin_client, word_factory):
    # Arrange: 注册Mock响应
    mock_addin_client.register_response("event_name", lambda req: {...})
    
    # Act: 执行动作
    result = await workspace.execute(action)
    
    # Assert: 验证结果
    assert result.success is True
```

## 目录结构

```
tests/contract_tests/
├── conftest.py           # Fixtures: workspace, mock_addin_client
├── mock_addin/
│   ├── client.py         # MockAddInClient实现
│   └── response_registry.py
├── factories/
│   └── word_factories.py # 数据工厂(mimesis/polyfactory)
└── word/
    └── test_*.py         # Contract测试
```

## 运行命令

```bash
poe test-unit        # 单元测试
poe test-integration # 集成测试
poe test-contract    # 契约测试
poe test-ci          # CI全量测试
```

## 覆盖率目标

- DTO层: ≥90%
- Service层: ≥80%
- Namespace层: ≥70%
- 整体: ≥75%

---

## Manual Tests (手动测试)

手动测试用于验证真实 Office 环境下的端到端功能，覆盖各种参数排列组合。

### 测试目录结构

```
manual_tests/
├── MANUAL_TEST.md                    # 总体指南
├── test_word_e2e.py                  # 快速验证
├── test_workspace_startup.py         # 启动测试
│
├── insert_text_e2e/                  # word:insert:text 参数组合
│   ├── README.md                     # ⭐ 测试说明文档（必需）
│   ├── test_basic_insert.py          # 基础插入 (4 tests)
│   ├── test_location_insert.py       # 位置参数 (4 tests)
│   └── test_format_insert.py         # 格式参数 (6 tests)
│
├── get_selected_content_e2e/         # word:get:selectedContent 参数组合
│   ├── README.md                     # ⭐ 测试说明文档（必需）
│   ├── test_basic_get.py             # 基础获取 (4 tests)
│   ├── test_options_get.py           # 选项参数 (5 tests)
│   └── test_edge_cases.py            # 边界情况 (4 tests)
│
├── replace_selection_e2e/            # word:replace:selection 参数组合
│   ├── README.md                     # ⭐ 测试说明文档（必需）
│   ├── test_text_replace.py          # 文本替换 (4 tests)
│   ├── test_format_replace.py        # 格式替换 (4 tests)
│   └── test_edge_cases.py            # 边界情况 (3 tests)
│
├── get_styles_e2e/                   # word:get:styles 参数组合
│   ├── README.md                     # ⭐ 测试说明文档（必需）
│   └── test_styles.py                # 样式获取 (5 tests)
│
└── connection_e2e/                   # 连接稳定性测试
    ├── README.md                     # ⭐ 测试说明文档（必需）
    ├── test_reconnection.py          # 重连测试
    └── test_multi_document.py        # 多文档测试
```

### README.md 文档（必需）

**作用：**
每个手动测试目录必须包含 `README.md` 文档，用于说明测试目的、运行方法和验证要点。

**必要性：**
- **开发协作**：帮助团队成员快速理解测试内容
- **使用指南**：提供清晰的运行命令和前置条件
- **验证标准**：明确预期结果和验证要点
- **维护文档**：记录测试变更历史和最后更新时间

**内容模板（参考兄弟目录）：**

```markdown
# <Event Name> E2E Tests

`word:<event_name>` 事件的手动端到端测试。

## 测试场景

| 测试编号 | 测试名称 | options/参数 | 描述 |
|---------|---------|-------------|------|
| 1 | 测试名称1 | `{param: value}` | 测试描述 |
| 2 | 测试名称2 | `{param: value}` | 测试描述 |

## 运行方式

### 前置条件
1. 启动 Workspace 服务器（测试会自动启动）
2. 在 Word 中加载 office-editor4ai Add-In
3. 确保文档已打开并连接

### 运行单个测试
```bash
uv run python manual_tests/<test_dir>/test_xxx.py --test 1
```

### 运行所有测试
```bash
uv run python manual_tests/<test_dir>/test_xxx.py --test all
```

## 验证要点

### 功能验证
- [ ] 验证点1
- [ ] 验证点2

### 数据验证
- [ ] 数据验证点1
- [ ] 数据验证点2

## 相关文档

- [Confluence 文档](链接) - 事件规范
- [Socket.IO API 标准](链接) - 完整协议规范

## 最后更新

YYYY-MM-DD
```

**参考示例：**
- `manual_tests/insert_text_e2e/README.md` - 详细格式（推荐）
- `manual_tests/get_styles_e2e/README.md` - 简洁格式


### 运行方式

```bash
# 从项目根目录运行单个测试
uv run python manual_tests/<test_dir>/test_xxx.py --test 1

# 运行某目录全部测试
uv run python manual_tests/<test_dir>/test_xxx.py --test all

# 示例：运行 insert_text 基础测试
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test all
```

### 参数组合矩阵

详细的测试用例设计、参数组合说明和代码模板见：
- **[reference.md](./reference.md)** - 手动测试参考文档
