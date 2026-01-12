# Word Get Document Structure End-to-End Tests

测试 `word:get:documentStructure` 事件的各种场景。

## 目录结构

```
get_document_structure_e2e/
├── __init__.py                 # 包初始化文件
├── test_basic_structure.py     # 基本结构获取测试（4个测试）
└── README.md                   # 本文档
```

## 测试概览

### 1. 基本结构获取测试 (`test_basic_structure.py`)

测试不同文档类型的结构信息获取。

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 空白文档结构 | 获取空白文档的结构信息 |
| 2 | 简单文档结构 | 获取包含简单文本的文档结构 |
| 3 | 复杂文档结构 | 获取包含多种元素的文档结构 |
| 4 | 多段落文档结构 | 获取包含大量段落的文档结构 |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 1
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 2
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 3
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 4

# 运行所有测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/get_document_structure_e2e

# 运行单个测试
uv run python test_basic_structure.py --test 1
uv run python test_basic_structure.py --test 2
uv run python test_basic_structure.py --test 3
uv run python test_basic_structure.py --test 4

# 运行所有测试
uv run python test_basic_structure.py --test all
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
   - 根据测试场景打开对应的 Word 文档（.docx）
   - 确保文档未被锁定

4. ✅ **文档准备**
   - 测试 1: 使用空白文档
   - 测试 2: 使用包含简单文本的文档
   - 测试 3: 使用包含表格、图片的复杂文档
   - 测试 4: 使用包含多个段落的长文档

---

## 测试流程

每个测试都会执行以下步骤：

```
1. 启动 Office Workspace (127.0.0.1:3000)
2. 等待 Word Add-In 连接 (最多30秒)
   ⚠️  如果超时，请检查 Add-In 是否正确加载
3. 获取已连接文档列表
4. 执行获取文档结构动作
5. 验证返回结果（段落数、表格数、图片数、节数）
6. 清理资源（停止 Workspace）
```

---

## 快速开始

### 1. 运行单个测试

**从项目根目录运行（推荐）：**
```bash
# 运行空白文档结构测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 1

# 运行简单文档结构测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 2

# 运行复杂文档结构测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 3

# 运行多段落文档结构测试
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test 4
```

**或者进入测试目录后运行：**
```bash
# 进入测试目录
cd manual_tests/get_document_structure_e2e

# 运行空白文档结构测试
uv run python test_basic_structure.py --test 1

# 运行简单文档结构测试
uv run python test_basic_structure.py --test 2

# 运行复杂文档结构测试
uv run python test_basic_structure.py --test 3

# 运行多段落文档结构测试
uv run python test_basic_structure.py --test 4
```

### 2. 运行所有测试

**从项目根目录运行（推荐）：**
```bash
# 运行所有测试（4个）
uv run python manual_tests/get_document_structure_e2e/test_basic_structure.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/get_document_structure_e2e

# 运行所有测试（4个）
uv run python test_basic_structure.py --test all
```

---

## 预期结果

### ✅ 成功情况

```
✅ Workspace 启动成功
✅ Add-In 已连接
✅ 使用文档: file:///path/to/document.docx
📝 获取文档结构...
✅ 获取成功
   段落数: 10
   表格数: 2
   图片数: 1
   节数: 1
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

3. **获取失败**
   ```
   ❌ 获取失败: {error details}
   ```
   解决方法：检查 Word 文档是否可访问，是否已损坏

---

## 测试场景说明

### 测试 1: 空白文档结构

**目的：** 验证对空白文档的结构统计是否正确。

**测试内容：**
- 使用完全空白的 Word 文档
- 验证 paragraphCount = 0 或 1（取决于是否包含空段落）
- 验证 tableCount = 0
- 验证 imageCount = 0
- 验证 sectionCount = 1

**预期结果：** 返回的结构信息应该反映空白文档的状态。

---

### 测试 2: 简单文档结构

**目的：** 验证对包含简单文本的文档的结构统计。

**测试内容：**
- 使用包含几段简单文本的文档
- 验证 paragraphCount > 0
- 验证其他计数正确

**预期结果：** 返回的段落数量应该与实际段落数一致。

---

### 测试 3: 复杂文档结构

**目的：** 验证对包含多种元素的文档的结构统计。

**测试内容：**
- 使用包含文本、表格、图片的复杂文档
- 验证 paragraphCount > 0
- 验证 tableCount > 0
- 验证 imageCount > 0

**预期结果：** 返回的所有计数应该与文档中实际元素数量一致。

---

### 测试 4: 多段落文档结构

**目的：** 验证对包含大量段落的文档的结构统计（性能测试）。

**测试内容：**
- 使用包含大量段落（例如 50+）的文档
- 验证 paragraphCount 的准确性

**预期结果：** 应该快速返回准确的段落计数。

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

### Q3: 计数不准确

**问题：** 返回的计数与实际不符

**解决方法：**
1. 手动检查文档中的元素数量
2. 确认文档中没有隐藏的元素
3. 检查段落计数是否包含空段落
4. 查看图片计数是否包含嵌入式和浮动图片

---

## 调试技巧

### 1. 查看详细日志

测试会输出详细的执行日志，包括：
- Workspace 启动状态
- Add-In 连接状态
- 文档列表
- 执行的动作
- 返回的结构信息

### 2. 手动验证

测试完成后，在 Word 中手动检查：
- 段落数是否正确
- 表格数是否正确
- 图片数是否正确
- 节数是否正确

### 3. 对比验证

可以使用 Word 的"字数统计"功能对比段落数量。

---

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/30769153) - 事件规范
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py) - 数据结构定义
- [Namespace 实现](../../office4ai/environment/workspace/socketio/namespaces/word.py) - 事件处理实现
- [单元测试](../../tests/unit_tests/office4ai/environment/workspace/socketio/namespaces/test_word_namespace.py) - 单元测试
- [契约测试](../../tests/contract_tests/word/test_get_document_structure.py) - 契约测试

---

## 最后更新

2026-01-12
