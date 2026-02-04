---
name: e2e-test-design
description: E2E 测试设计指南，包含双重验证机制（协议验证 + 文档内容验证）
usage: |-
  设计 E2E 测试时参考此指南，特别是涉及文档内容修改的测试。
  区分只读事件和修改事件，选择合适的验证策略。
---

# E2E Test Design Skill

## 核心概念：双重验证机制

E2E 测试支持两种验证模式：

```
协议验证 (DataValidator)    → 验证服务端返回的 data 结构和值
文档验证 (ContentValidator) → 通过 python-docx 验证实际文档内容
```

### 为什么需要双重验证？

- 服务端可能返回 `success: true`，但文档未被正确修改
- 仅依赖协议返回不够鲁棒，无法验证修改后的具体内容
- 测试可能因 Add-In 实现缺陷而通过，但实际功能有问题

---

## 事件分类

### 只读事件（使用 DataValidator）

这些事件不修改文档，只需验证协议返回：

| 事件 | 说明 |
|------|------|
| `word:get:documentStats` | 获取文档统计 |
| `word:get:documentStructure` | 获取文档结构 |
| `word:get:selectedContent` | 获取选中内容 |
| `word:get:styles` | 获取样式列表 |

**验证策略**：
```python
validator=lambda data: data.get("wordCount", 0) > 0
```

### 修改事件（使用 ContentValidator）

这些事件会修改文档，应使用双重验证：

| 事件 | 说明 |
|------|------|
| `word:insert:text` | 插入文本 |
| `word:replace:selection` | 替换选中内容 |
| `word:insert:table` | 插入表格 |
| `word:set:style` | 设置样式 |

**验证策略**：
```python
validator=lambda data, reader: (
    data.get("success", False) and
    reader.contains("预期插入的文本")
)
```

---

## Validator 类型定义

```python
from manual_tests.e2e_base import (
    DataValidator,      # Callable[[dict[str, Any]], bool]
    ContentValidator,   # Callable[[dict[str, Any], DocumentReader], bool]
    Validator,          # DataValidator | ContentValidator
    DocumentReader,     # 文档内容读取器
)
```

### DocumentReader API

```python
@dataclass
class DocumentReader:
    path: Path

    @property
    def doc(self) -> Document:
        """懒加载 python-docx Document 对象"""

    @property
    def paragraphs(self) -> list[str]:
        """获取所有段落文本"""

    @property
    def text(self) -> str:
        """获取全文文本（段落用换行连接）"""

    @property
    def table_count(self) -> int:
        """获取表格数量"""

    def reload(self) -> None:
        """重新加载文档（修改后调用）"""

    def contains(self, text: str) -> bool:
        """检查文档是否包含指定文本"""

    def paragraph_contains(self, index: int, text: str) -> bool:
        """检查指定段落是否包含文本"""

    def get_paragraph(self, index: int) -> str | None:
        """获取指定段落的文本"""
```

---

## 验证示例

### 示例 1：只读事件（DataValidator）

```python
TestCase(
    name="空白文档统计",
    fixture_name="empty.docx",
    description="空白文档应有 0 字",
    expected=ExpectedStats(word_count=0),
    # 传统验证器：仅检查协议返回
    validator=lambda data: data.get("wordCount", 0) == 0,
)
```

### 示例 2：修改事件（ContentValidator）

```python
TestCase(
    name="插入文本测试",
    fixture_name="empty.docx",
    description="向空白文档插入 'Hello World'",
    # 双重验证器：检查协议返回 + 文档内容
    validator=lambda data, reader: (
        # 协议验证
        data.get("success", False) and
        # 文档内容验证
        reader.contains("Hello World")
    ),
)
```

### 示例 3：复杂内容验证

```python
def validate_table_insert(data: dict, reader: DocumentReader) -> bool:
    """验证表格插入"""
    # 重新加载以获取最新内容
    reader.reload()

    # 协议验证
    if not data.get("success", False):
        return False

    # 文档验证：检查表格数量
    if reader.table_count < 1:
        return False

    # 检查表格内容
    table = reader.doc.tables[0]
    return table.rows[0].cells[0].text == "预期内容"

TestCase(
    name="插入表格测试",
    fixture_name="empty.docx",
    validator=validate_table_insert,
)
```

---

## 最佳实践

### 1. 选择合适的验证模式

```
只读事件 → DataValidator（简单 lambda）
修改事件 → ContentValidator（双重验证）
```

### 2. 验证顺序

```python
def my_validator(data, reader):
    # 1. 先验证协议返回
    if not data.get("success"):
        return False

    # 2. 再验证文档内容
    reader.reload()  # 确保获取最新内容
    return reader.contains("预期内容")
```

### 3. 避免过度验证

- 只验证与测试目标相关的内容
- 不要验证无关的副作用
- 保持验证逻辑简洁

### 4. 处理异步修改

某些修改可能需要时间生效，必要时添加等待：

```python
import asyncio

async def test_with_delay():
    result = await workspace.execute(action)
    await asyncio.sleep(0.5)  # 等待文档保存

    reader = DocumentReader(fixture.working_path)
    assert reader.contains("预期内容")
```

---

## TestCase 定义模板

```python
from dataclasses import dataclass, field
from typing import Any
from manual_tests.e2e_base import ExpectedStats, Validator

@dataclass
class TestCase:
    name: str
    fixture_name: str
    description: str
    expected: ExpectedStats | None = None
    validator: Validator | None = None  # DataValidator | ContentValidator
    tags: list[str] = field(default_factory=list)
```

---

## 相关文件

| 文件 | 说明 |
|------|------|
| `manual_tests/e2e_base.py` | DocumentReader、Validator 类型定义 |
| `manual_tests/get_document_structure_e2e/` | 只读事件测试示例 |
| `manual_tests/get_document_stats_e2e/` | 只读事件测试示例 |

---

## 运行验证

```bash
# 类型检查
poe typecheck

# 运行 E2E 测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test all
```

---

**最后更新**: 2026-02-04
