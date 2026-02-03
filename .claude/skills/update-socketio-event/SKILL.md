---
description: 更新现有的 Socket.IO 事件。分析协议变更、更新 DTO/Namespace、同步测试。
argument-hint: <event_name> 例如 word:get:documentContent
---

# Socket.IO 事件更新 Skill

根据事件名更新现有的 Socket.IO 事件实现，遵循项目既有模式。

## 核心思路

### 为什么更新比创建更复杂

更新事件需要处理三个额外维度：

- **向后兼容性**：新增可选字段安全，新增必填字段/删除字段有破坏性
- **回归保护**：现有功能必须保持正常工作
- **废弃清理**：标记或移除不再使用的字段

### 涉及文件定位

```
DTO 定义      → office4ai/environment/workspace/dtos/<platform>.py
Namespace     → office4ai/environment/workspace/socketio/namespaces/<platform>.py
单元测试      → tests/unit_tests/office4ai/environment/workspace/dtos/
契约测试      → tests/contract_tests/word/
手动测试      → manual_tests/<event_name>_e2e/
```

## 执行步骤

### Step 1: 分析协议变更

对比新旧 OASP 协议规范，生成变更清单。

**查找新协议**：
1. **本地 oasp-protocol 目录**（推荐）：直接阅读对应 `.md` 文件
2. **在线文档**：https://a2c-smcp.github.io/oasp-protocol/latest/
   - 如用户指定版本，使用对应版本 URL

**变更清单模板**：

| 变更类型 | 字段/行为 | 旧值 | 新值 | 影响范围 |
|---------|----------|-----|-----|---------|
| 新增 | `options.maxLength` | - | `int \| None` | DTO, 测试 |
| 修改 | `format.color` | `str` | `str \| None` | DTO, 测试 |
| 废弃 | `legacyParam` | 存在 | 移除 | DTO, 测试 |
| 行为 | 空选区处理 | 返回错误 | 返回空数据 | Namespace, 测试 |

### Step 2: 更新 DTO

在对应平台的 DTO 文件中修改。

**参考模式**：[office4ai/environment/workspace/dtos/word.py](office4ai/environment/workspace/dtos/word.py)

**字段变更模式**：

```python
# 新增字段：添加可选字段
new_field: str | None = Field(None, alias="newField", description="...")

# 废弃字段：标记 deprecated（保持兼容）或直接删除（破坏性）
legacy_field: str | None = Field(None, alias="legacyField", deprecated=True)
```

**向后兼容性原则**：
| 场景 | 风险 | 处理方式 |
|-----|-----|---------|
| 新增可选字段 | ✅ 安全 | 直接添加 |
| 新增必填字段 | ⚠️ 破坏性 | 需版本协调 |
| 修改字段类型 | ⚠️ 评估 | 检查现有用法 |
| 删除字段 | ⚠️ 破坏性 | 先标记 deprecated |

### Step 3: 更新 Namespace（如有行为变更）

在对应平台的 Namespace 文件中修改处理逻辑。

**参考模式**：[office4ai/environment/workspace/socketio/namespaces/word.py](office4ai/environment/workspace/socketio/namespaces/word.py)

### Step 4: 更新单元测试

更新 DTO 测试，确保新功能和回归保护。

**参考模式**：[tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py](tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py)

**测试更新策略**：
| 变更类型 | 测试策略 |
|---------|---------|
| 新功能 | 添加新测试用例覆盖新字段/行为 |
| 旧功能 | 确保现有测试仍然通过（回归保护） |
| 废弃点 | 移除相关测试或标记跳过 |

### Step 5: 更新手动测试（如有）

更新 `manual_tests/<event_name>_e2e/` 目录。

**需要更新的内容**：
- `README.md`：添加新测试场景、标记废弃场景、更新日期
- 测试文件：新增测试函数、移除废弃参数测试、更新 `TEST_MAPPING`

**详细指南**：参考 [test-checklist Skill](.claude/skills/test-checklist/SKILL.md) 中的 Manual Tests 章节。

### Step 6: 验证

```bash
poe lint          # 代码检查
poe typecheck     # 类型检查
poe test-unit     # 单元测试 - 验证 DTO 变更
poe test-contract # 契约测试 - 验证协议兼容性
```

手动测试（如有更新）：
```bash
uv run python manual_tests/<event_name>_e2e/test_xxx.py --test all
```

## 完成检查清单

- [ ] 协议变更已分析并记录变更清单
- [ ] DTO 已更新（新增/修改/废弃）
- [ ] Namespace 已更新（如有行为变更）
- [ ] 单元测试已更新并通过
- [ ] 手动测试已更新（如适用）
- [ ] 所有验证命令通过

## 关键类参考

- [BaseRequest](office4ai/environment/workspace/dtos/common.py) - 请求基类
- [SocketIOBaseModel](office4ai/environment/workspace/dtos/common.py) - 嵌套类型基类
- [ErrorCode](office4ai/environment/workspace/dtos/common.py) - 错误码常量
