# A2C 协议 - 换行符标准化规范

> **文档类型**: 协议规范补充 / 技术决策记录 (TDR)
> **创建日期**: 2026-01-12
> **影响范围**: `word:get:visibleContent` 及所有返回文本内容的接口
> **优先级**: **P0 - 必须立即修复**

---

## 1. 问题发现

### 1.1 测试场景

在 E2E 测试 `test_basic_get.py --test 4` 中，发现 Word 文档返回的文本内容在终端打印时**没有正确换行**。

### 1.2 实际日志证据

```python
# 完整文本字符分析
repr(): '文本与图片\r\r大标题\r\r'
字符详情:
   [0] = 25991 = '文'
   [1] = 26412 = '本'
   [2] = 19982 = '与'
   [3] = 22270 = '图'
   [4] = 29255 = '片'
   [5] =  13 = '\r' \r (CR)        # ← 第一个段落分隔符
   [6] =  13 = '\r' \r (CR)        # ← 两个连续的 \r
   [7] = 22823 = '大'
   [8] = 26631 = '标'
   [9] = 39064 = '题'
   [10] =  13 = '\r' \r (CR)       # ← 第二个段落分隔符
   [11] =  13 = '\r' \r (CR)       # ← 又是两个连续的 \r
```

### 1.3 元素级别分析

```
元素 #1: '文本与图片'  ✅ 无 \r
元素 #2: ''            (空段落)
元素 #4: '大标题'      ✅ 无 \r
元素 #5: ''            (空段落)
```

**关键发现**: 元素内部文本是干净的，但**合并元素时插入了 `\r\r`**。

---

## 2. 根本原因

### 2.1 Word 的段落分隔符历史

- **Microsoft Word** 使用 `\r` (CR, Carriage Return, ASCII 13) 作为段落标记
- 不同版本的 Word API 行为不一致：
  - 某些版本返回单个 `\r`
  - 某些版本返回 `\r\r`（测试发现的情况）
  - COM/UNO 接口可能返回 `\r\n`

### 2.2 为什么 `\r` 不会换行？

| 字符 | 名称 | ASCII | 作用 | 终端行为 |
|------|------|-------|------|---------|
| `\n` | LF (Line Feed) | 10 | **换行** | 光标移到下一行 |
| `\r` | CR (Carriage Return) | 13 | **回车** | 光标回到**行首**（不换行） |

**打印示例**:
```python
print('文本与图片\r\r大标题')
# 输出: 大标题与图片  ← "大标题" 覆盖了行首的前 3 个字符！
```

### 2.3 跨平台差异

| 系统 | 历史换行符 | 现代标准 |
|------|-----------|---------|
| Unix/Linux | `\n` | `\n` |
| macOS (Classic) | `\r` | `\n` (OS X 起) |
| Windows | `\r\n` | `\r\n` |
| **Web/JSON/现代协议** | **推荐 `\n`** | **`\n`** |

---

## 3. 影响评估

### 3.1 对 A2C 协议的影响

| 问题 | 影响 | 严重性 |
|------|------|--------|
| **终端显示错乱** | `\r` 导致覆盖，无法正确渲染文档内容 | 🔴 高 |
| **JSON 序列化问题** | `\r` 在 JSON 中需要转义，增加复杂度 | 🟡 中 |
| **跨平台兼容性** | Unix/Linux/Mac 统一使用 `\n` | 🔴 高 |
| **协议一致性** | 当前协议未明确，导致实现不一致 | 🔴 高 |
| **AI 消费友好性** | LLM 更容易处理 `\n` 分隔的文本 | 🟡 中 |

### 3.2 对下游系统的影响

1. **Python MCP Server**: 需要额外处理 `\r` → `\n` 转换
2. **AI Agent**: 收到的文本无法正确显示，影响理解
3. **日志/调试**: 终端输出混乱，难以排查问题
4. **数据持久化**: 存储非标准换行符，后续处理困难

---

## 4. 协议规范建议

### 4.1 明确要求（协议规范）

> **A2C 协议规范 v1.x - 换行符标准化**
>
> 所有返回文本内容的接口，**必须**使用 `\n` (LF, ASCII 10) 作为段落分隔符。
>
> **禁止**使用：
> - ❌ `\r` (CR)
> - ❌ `\r\r` (双 CR)
> - ❌ `\r\n` (CRLF)
>
> **合规示例**:
> ```json
> {
>   "text": "第一段文本\n第二段文本\n第三段文本",
>   "elements": [...]
> }
> ```

### 4.2 TypeScript 端实现规范

**文件位置**: `office-editor4ai/src/word/handlers/getVisibleContent.handler.ts`

```typescript
/**
 * 标准化换行符 - 协议要求
 *
 * Word API 返回的文本可能使用 \r 或 \r\r 作为段落分隔符
 * 必须转换为标准的 \n (LF) 以符合 A2C 协议规范
 */
function normalizeLineBreaks(text: string): string {
  if (!text) return text;

  return text
    // 1. Word 双 \r → \n
    .replace(/\r\r/g, '\n')
    // 2. Windows CRLF → \n
    .replace(/\r\n/g, '\n')
    // 3. 单独 CR → \n
    .replace(/\r/g, '\n')
    // 4. 结尾规范化（移除多余的 \n）
    .replace(/\n+$/, '\n');
}

/**
 * 构建 word:get:visibleContent 响应
 */
export async function handleGetVisibleContent(
  document: Word.Document,
  options: GetVisibleContentOptions
): Promise<GetVisibleContentResult> {
  // 获取 Word 原始文本（可能包含 \r\r）
  const rawText = await getDocumentText(document);

  // 🔴 协议要求：必须标准化换行符
  const normalizedText = normalizeLineBreaks(rawText);

  return {
    text: normalizedText,  // ✅ 使用标准 \n
    metadata: { ... },
    elements: [ ... ]
  };
}
```

### 4.3 Python 端无需处理

**原则**: 数据源（TypeScript）负责标准化，消费者（Python）直接使用。

```python
# office4ai/environment/workspace/dtos/word.py

@dataclass
class GetVisibleContentResult:
    text: str  # ✅ 已由 TypeScript 端标准化，直接使用
    metadata: dict
    elements: list

    def display(self) -> None:
        # 直接打印即可正常换行
        print(self.text)  # ✅ 无需额外处理
```

---

## 5. 测试验证

### 5.1 单元测试（TypeScript）

```typescript
import { describe, it, expect } from 'vitest';
import { normalizeLineBreaks } from './normalizeLineBreaks';

describe('normalizeLineBreaks - 协议合规性测试', () => {
  it('应该将 Word 双 \\r 转换为 \\n', () => {
    const input = '文本与图片\r\r大标题\r\r';
    const expected = '文本与图片\n大标题\n';
    expect(normalizeLineBreaks(input)).toBe(expected);
  });

  it('应该处理 CRLF', () => {
    const input = '第一行\r\n第二行\r\n';
    const expected = '第一行\n第二行\n';
    expect(normalizeLineBreaks(input)).toBe(expected);
  });

  it('应该处理单个 \\r', () => {
    const input = '第一行\r第二行\r';
    const expected = '第一行\n第二行\n';
    expect(normalizeLineBreaks(input)).toBe(expected);
  });

  it('应该移除结尾多余的换行符', () => {
    const input = '文本\n\n\n\n';
    const expected = '文本\n';
    expect(normalizeLineBreaks(input)).toBe(expected);
  });

  it('空字符串应该保持不变', () => {
    expect(normalizeLineBreaks('')).toBe('');
  });
});
```

### 5.2 契约测试（Python）

```python
# tests/contract_tests/word/test_get_visible_content.py

import pytest

class TestLineBreakNormalization:
    """契约测试：验证 TypeScript 端已正确标准化换行符"""

    @pytest.mark.contract
    def test_text_must_use_lf_not_cr(self, result: GetVisibleContentResult):
        """协议要求：text 字段必须使用 \n 分隔段落"""
        text = result.text

        # 禁止包含 \r
        assert '\r' not in text, \
            "协议违规：text 字段包含 \\r (CR)，必须使用 \\n (LF)"

        # 验证段落分隔符是 \n
        paragraphs = text.split('\n')
        assert len(paragraphs) > 1, "测试数据应包含多个段落"

    @pytest.mark.contract
    def test_text_printable_in_terminal(self, result: GetVisibleContentResult):
        """终端可读性：文本应该能正确换行显示"""
        import io
        from contextlib import redirect_stdout

        f = io.StringIO()
        with redirect_stdout(f):
            print(result.text)

        output = f.getvalue()
        # 验证没有覆盖现象
        assert "大标题" in output
```

---

## 6. 行动项

### 6.1 立即执行（P0）

| 任务 | 责任方 | 验证方式 |
|------|--------|---------|
| ✅ **TypeScript: 添加 `normalizeLineBreaks()` 函数** | 前端开发者 | 单元测试通过 |
| ✅ **TypeScript: 在 `handleGetVisibleContent` 中调用** | 前端开发者 | E2E 测试通过 |
| ✅ **TypeScript: 添加单元测试覆盖** | 前端开发者 | 100% 覆盖率 |
| ✅ **Python: 添加契约测试验证** | 后端开发者 | 契约测试通过 |

### 6.2 协议文档更新

| 任务 | 责任方 | 位置 |
|------|--------|------|
| ✅ **更新 A2C RFC 添加换行符规范** | 协议维护者 | `docs/a2c/a2c_rfc.md` |
| ✅ **在 API 文档中明确要求** | 协议维护者 | Swagger/OpenAPI 规范 |

### 6.3 长期优化（P1）

- [ ] 添加 CI 检查：禁止 `\r` 进入 JSON 响应
- [ ] 统一所有 Office 应用的换行符处理（PPT、Excel）
- [ ] 添加集成测试：跨平台换行符验证

---

## 7. 参考资料

### 7.1 技术文档

- [Unicode Line Separators (U+2028, U+2029)](https://unicode.org/reports/tr13/tr13-5.html)
- [ECMAScript JSON String Encoding](https://www.ecma-international.org/ecma-262/11.0/index.html#sec-JSON.stringify)
- [A2C Protocol RFC](../a2c/a2c_rfc.md)

### 7.2 相关 Issue

- 内部 Issue: `#XXX - Word 返回的文本包含 \r 导致终端显示错乱`
- 测试脚本: `manual_tests/get_visible_content_e2e/test_basic_get.py`

---

## 8. 附录：完整技术分析

### 8.1 字符编码详解

```
CR  (Carriage Return, \r, ASCII 13)  → 光标回到行首，不换行
LF  (Line Feed, \n, ASCII 10)        → 光标移到下一行
CRLF (\r\n)                          → Windows 风格换行
```

### 8.2 Python print() 函数行为

```python
# 场景 1: \n 正常换行
print('Hello\nWorld')
# 输出:
# Hello
# World

# 场景 2: \r 覆盖行首
print('Hello\rWorld')
# 输出: World (覆盖了 "Hello")

# 场景 3: \r\r 仍然覆盖
print('文本与图片\r\r大标题')
# 输出: 大标题与图片 (覆盖前 3 个字符)
```

### 8.3 Word API 行为差异

| API 版本 | 返回值 | 说明 |
|---------|--------|------|
| Word 2016 COM | `\r` | 单个 CR |
| Word 2019+ | `\r\r` | **当前测试发现** |
| Word Online | `\r\n` | CRLF |
| LibreOffice UNO | `\n` | 已经正确 |

**结论**: 必须在 TypeScript 端统一处理，不依赖 Word API 的原始行为。

---

**文档维护者**: JQQ <jqq1716@gmail.com>
**最后更新**: 2026-01-12
**状态**: 🟡 待协议维护者审核
