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
