# Word Replace Text End-to-End Tests

测试 `word:replace:text` 事件的各种入参组合。

Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30801921

## 目录结构

```
replace_text_e2e/
├── __init__.py                 # 包初始化文件
├── test_basic_replace.py       # 基础替换测试（6个测试）
├── test_options_replace.py     # 选项参数测试（5个测试）
├── test_edge_cases.py          # 边界情况测试（4个测试）
└── README.md                   # 本文档
```

## 测试概览

### 1. 基础替换测试 (`test_basic_replace.py`)

测试基础的文本查找和替换功能。

| 测试编号 | 测试名称 | searchText | replaceText | 描述 |
|---------|---------|-----------|-------------|------|
| 1 | 简单文本替换（全部） | "old" | "new" | 替换文档中所有匹配项 |
| 2 | 简单文本替换（首个） | "test" | "exam" | 仅替换第一个匹配项 |
| 3 | 替换为空（删除） | "delete" | "" | 删除匹配的文本 |
| 4 | 多行文本替换 | "line1\nline2" | "new\\ncontent" | 替换多行文本 |
| 5 | 特殊字符替换 | "Café" | "Coffee" | 替换含特殊字符的文本 |
| 6 | 长文本替换 | 长文本 | 长文本 | 替换较长段落 |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test 1
uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test 2
# ... 其他测试

# 运行所有测试
uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test all
```

**方式二：进入测试目录后运行**
```bash
cd manual_tests/replace_text_e2e

# 运行单个测试
uv run python test_basic_replace.py --test 1

# 运行所有测试
uv run python test_basic_replace.py --test all
```

---

### 2. 选项参数测试 (`test_options_replace.py`)

测试 `options` 参数的各种组合。

| 测试编号 | 测试名称 | options | 描述 |
|---------|---------|---------|------|
| 1 | 大小写敏感替换 | `{matchCase: true}` | 区分大小写替换 |
| 2 | 大小写不敏感替换 | `{matchCase: false}` | 不区分大小写替换 |
| 3 | 全字匹配替换 | `{matchWholeWord: true}` | 仅匹配完整单词 |
| 4 | 全字匹配+全部替换 | `{matchWholeWord: true, replaceAll: true}` | 全字匹配并替换所有 |
| 5 | 组合选项 | `{matchCase: true, matchWholeWord: true, replaceAll: true}` | 所有选项启用 |

**运行方式：**

**从项目根目录运行（推荐）：**
```bash
# 运行单个测试
uv run python manual_tests/replace_text_e2e/test_options_replace.py --test 1

# 运行所有测试
uv run python manual_tests/replace_text_e2e/test_options_replace.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/replace_text_e2e

# 运行单个测试
uv run python test_options_replace.py --test 1

# 运行所有测试
uv run python test_options_replace.py --test all
```

---

### 3. 边界情况测试 (`test_edge_cases.py`)

测试 `word:replace:text` 的边界情况和错误处理。

| 测试编号 | 测试名称 | 描述 | 预期结果 |
|---------|---------|------|---------|
| 1 | 空搜索文本 | searchText 为空字符串 | 返回错误码 4001 (MISSING_PARAM) |
| 2 | 空替换文本 | replaceText 为空字符串 | 返回错误码 4001 (MISSING_PARAM) |
| 3 | 无匹配项 | 搜索不存在的文本 | 返回 replaceCount=0 |
| 4 | 大量替换 | 文档中存在数百个匹配项 | 返回正确的 replaceCount |

**运行方式：**

**从项目根目录运行（推荐）：**
```bash
# 运行单个测试
uv run python manual_tests/replace_text_e2e/test_edge_cases.py --test 1

# 运行所有测试
uv run python manual_tests/replace_text_e2e/test_edge_cases.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/replace_text_e2e

# 运行单个测试
uv run python test_edge_cases.py --test 1

# 运行所有测试
uv run python test_edge_cases.py --test all
```

---

## 前置条件

运行这些测试前，请确保：

1. ✅ **Workspace 服务器可访问**
   - 测试会自动启动 Workspace 在 `http://127.0.0.1:3000`
   - 无需手动启动服务器

2. ✅ **Word Add-In 已加载**
   - 在 Word 中加载 `office-editor4ai` Add-In
   - Add-In 会自动连接到 `http://127.0.0.1:3000`

3. ✅ **Word 文档已打开**
   - 打开任意 Word 文档（.docx）
   - 确保文档未被锁定

4. ✅ **文档准备**
   - 在文档中预先输入测试所需的文本
   - 根据测试场景准备不同的文档内容
   - 建议创建测试文本便于验证替换结果

---

## 测试流程

每个测试都会执行以下步骤：

```
1. 启动 Office Workspace (127.0.0.1:3000)
2. 等待 Word Add-In 连接 (最多30秒)
   ⚠️  如果超时，请检查 Add-In 是否正确加载
3. 获取已连接文档列表
4. 执行文本查找和替换动作
5. 验证返回结果
6. 清理资源（停止 Workspace）
```

---

## 快速开始

### 1. 运行单个测试

**从项目根目录运行（推荐）：**
```bash
# 运行基础替换测试 1
uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test 1

# 运行选项参数测试 1
uv run python manual_tests/replace_text_e2e/test_options_replace.py --test 1

# 运行边界情况测试 1
uv run python manual_tests/replace_text_e2e/test_edge_cases.py --test 1
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/replace_text_e2e

# 运行基础替换测试 1
uv run python test_basic_replace.py --test 1

# 运行选项参数测试 1
uv run python test_options_replace.py --test 1

# 运行边界情况测试 1
uv run python test_edge_cases.py --test 1
```

### 2. 运行所有测试

**从项目根目录运行（推荐）：**
```bash
# 运行所有基础替换测试（6个）
uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test all

# 运行所有选项参数测试（5个）
uv run python manual_tests/replace_text_e2e/test_options_replace.py --test all

# 运行所有边界情况测试（4个）
uv run python manual_tests/replace_text_e2e/test_edge_cases.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/replace_text_e2e

# 运行所有基础替换测试
uv run python test_basic_replace.py --test all

# 运行所有选项参数测试
uv run python test_options_replace.py --test all

# 运行所有边界情况测试
uv run python test_edge_cases.py --test all
```

### 3. 完整测试流程

如果你想测试所有场景，依次运行：

**从项目根目录运行（推荐）：**
```bash
# 1. 基础替换
uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test all

# 2. 选项参数
uv run python manual_tests/replace_text_e2e/test_options_replace.py --test all

# 3. 边界情况
uv run python manual_tests/replace_text_e2e/test_edge_cases.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/replace_text_e2e

# 1. 基础替换
uv run python test_basic_replace.py --test all

# 2. 选项参数
uv run python test_options_replace.py --test all

# 3. 边界情况
uv run python test_edge_cases.py --test all
```

---

## 预期结果

### ✅ 成功情况

```
✅ Workspace 启动成功
✅ Add-In 已连接
✅ 使用文档: file:///path/to/document.docx
📝 查找文本: 'old'
📝 替换文本: 'new'
📝 选项: {'replaceAll': true}
✅ 替换成功
   返回数据: {'replaceCount': 5}
✅ 测试完成
```

### ❌ 失败情况

如果测试失败，请检查：

1. **Add-In 未连接**
   ```
   ❌ 超时：未检测到 Add-In 连接
   请检查:
   1. Word Add-In 是否已加载
   2. Add-In 是否能访问 http://127.0.0.1:3000
   3. 浏览器控制台是否有错误
   ```

2. **文档未找到**
   ```
   ❌ 未找到已连接文档
   ```
   解决方法：在 Word 中打开一个文档

3. **缺少必需参数**
   ```
   ❌ 替换失败: MISSING_PARAM
   错误码: 4001
   ```
   解决方法：确保 searchText 和 replaceText 都不为空

4. **文档未找到**
   ```
   ❌ 替换失败: DOCUMENT_NOT_FOUND
   错误码: 3001
   ```
   解决方法：确认文档路径正确且可访问

---

## 测试场景说明

### 基础替换测试

**目的：** 验证基本的文本查找和替换功能是否正常工作。

**测试内容：**
- 简单词替换（全部/首个）
- 删除文本（替换为空）
- 多行文本替换
- 特殊字符替换
- 长文本替换

**预期结果：** 文档中的文本被正确查找和替换，返回正确的 `replaceCount`。

**准备工作：**
1. 在 Word 文档中输入测试所需的文本
2. 根据测试场景准备不同的内容
3. 运行测试前确认文档中包含要搜索的文本

---

### 选项参数测试

**目的：** 验证 `options` 参数的各种组合是否正确工作。

**测试内容：**
- `matchCase` - 大小写敏感匹配
- `matchWholeWord` - 全字匹配
- `replaceAll` - 替换所有匹配项
- 组合选项测试

**预期结果：** 替换结果符合选项指定的规则。

**验证方法：** 在 Word 中手动检查替换结果，确认符合预期。

**准备工作：**
1. 在 Word 文档中输入包含大小写变体的文本
2. 准备包含部分匹配单词的文本
3. 根据测试场景准备相应的文档内容

---

### 边界情况测试

**目的：** 验证异常情况的处理。

**测试内容：**
- 空搜索文本（预期失败）
- 空替换文本（预期失败）
- 无匹配项（返回 0）
- 大量替换（性能测试）

**预期结果：**
- 测试 1-2：返回错误码 4001
- 测试 3：返回 replaceCount=0
- 测试 4：返回正确的 replaceCount

**准备工作：**
1. 测试 1-2：正常运行，测试会发送空字符串
2. 测试 3：确保文档中不包含搜索文本
3. 测试 4：准备包含大量重复文本的文档

---

## 常见问题

### Q1: Add-In 连接超时

**问题：** `❌ 超时：未检测到 Add-In 连接`

**解决方法：**
1. 确认 Word Add-In 已加载
2. 检查 Add-In 是否能访问 `http://127.0.0.1:3000`
3. 打开浏览器开发者工具查看控制台错误

### Q2: 文档未找到

**问题：** `❌ 未找到已连接文档`

**解决方法：**
1. 在 Word 中打开一个文档
2. 确保文档已完全加载
3. 重新运行测试

### Q3: 替换结果不符合预期

**问题：** 替换后的内容与预期不符

**解决方法：**
1. 确认文档中包含要搜索的文本
2. 检查 options 参数是否正确设置
3. 验证大小写和全字匹配选项
4. 重新运行测试

### Q4: replaceCount 不正确

**问题：** 返回的替换次数与实际不符

**解决方法：**
1. 手动在 Word 中统计匹配项
2. 检查选项参数是否影响匹配规则
3. 确认 replaceAll 参数是否正确设置

### Q5: 特殊字符替换失败

**问题：** 包含特殊字符的文本无法正确替换

**解决方法：**
1. 确认特殊字符的编码正确
2. 检查文本是否包含不可见字符
3. 尝试使用转义字符

---

## 调试技巧

### 1. 查看详细日志

测试会输出详细的执行日志，包括：
- Workspace 启动状态
- Add-In 连接状态
- 文档列表
- 执行的动作参数
- 返回的结果数据

### 2. 手动验证

测试完成后，在 Word 中手动检查：
- 查找的文本是否都被替换
- 替换后的内容是否正确
- 替换次数是否符合预期

### 3. 对比测试

运行多个测试后，对比不同参数的效果：
```bash
# 先运行默认选项的测试
uv run python test_basic_replace.py --test 1

# 再运行带选项的测试
uv run python test_options_replace.py --test 1

# 在 Word 中对比两次替换的差异
```

---

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/30801921) - word:replace:text 事件规范
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py) - Word DTOs
- [Namespace 实现](../../office4ai/environment/workspace/socketio/namespaces/word.py) - Word 命名空间
- [单元测试](../../tests/unit_tests/office4ai/environment/workspace/dtos/test_replace_text_dtos.py) - DTO 单元测试
- [契约测试](../../tests/contract_tests/word/test_replace_text.py) - 契约测试
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范

---

## 最后更新

2026-01-13
