# Word Replace Selection End-to-End Tests

测试 `word:replace:selection` 事件的各种入参组合。

## 目录结构

```
replace_selection_e2e/
├── __init__.py                 # 包初始化文件
├── test_text_replace.py        # 文本替换测试（4个测试）
├── test_format_replace.py      # 格式替换测试（4个测试）
├── test_edge_cases.py          # 边界情况测试（3个测试）
└── README.md                   # 本文档
```

## 测试概览

### 1. 文本替换测试 (`test_text_replace.py`)

测试基础的文本替换功能，使用默认参数。

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 替换为纯文本 | 替换选中的内容为 "Hello World" |
| 2 | 替换为多行文本 | 替换为包含换行符的文本 |
| 3 | 替换为特殊字符 | 替换为特殊字符文本 |
| 4 | 替换为长文本 | 替换为较长的文本段落（>100字符） |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test 1
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test 2
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test 3
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test 4

# 运行所有测试
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/word/replace_selection_e2e

# 运行单个测试
uv run python test_text_replace.py --test 1
uv run python test_text_replace.py --test 2
uv run python test_text_replace.py --test 3
uv run python test_text_replace.py --test 4

# 运行所有测试
uv run python test_text_replace.py --test all
```

---

### 2. 格式替换测试 (`test_format_replace.py`)

测试带格式参数的文本替换功能。

| 测试编号 | 测试名称 | format 参数 | 描述 |
|---------|---------|------------|------|
| 1 | 替换为粗体文本 | `{bold: true}` | 替换为粗体文本 |
| 2 | 替换为斜体文本 | `{italic: true}` | 替换为斜体文本 |
| 3 | 替换为带字体格式的文本 | `{fontName: "Arial", fontSize: 16, bold: true}` | 替换为指定字体和大小的文本 |
| 4 | 替换为带颜色和下划线的文本 | `{color: "#FF0000", underline: true, bold: true}` | 替换为红色下划线粗体文本 |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test 1
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test 2
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test 3
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test 4

# 运行所有测试
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/word/replace_selection_e2e

# 运行单个测试
uv run python test_format_replace.py --test 1
uv run python test_format_replace.py --test 2
uv run python test_format_replace.py --test 3
uv run python test_format_replace.py --test 4

# 运行所有测试
uv run python test_format_replace.py --test all
```

---

### 3. 边界情况测试 (`test_edge_cases.py`)

测试 `word:replace:selection` 的边界情况和错误处理。

| 测试编号 | 测试名称 | 描述 | 预期结果 |
|---------|---------|------|---------|
| 1 | 空选择替换 | 无选中内容时发送替换请求 | 返回错误码 3002 (SELECTION_EMPTY) |
| 2 | 替换为空字符串 | 删除选中的内容 | 返回 replaced=True, characterCount=0 |
| 3 | 替换为图片 | 替换为 base64 编码的图片 | 返回 replaced=True |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test 1
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test 2
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test 3

# 运行所有测试
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/word/replace_selection_e2e

# 运行单个测试
uv run python test_edge_cases.py --test 1
uv run python test_edge_cases.py --test 2
uv run python test_edge_cases.py --test 3

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
   - 对于替换测试，请先在文档中输入一些文本
   - 选中需要替换的文本内容
   - 建议创建测试文本便于验证替换结果

---

## 测试流程

每个测试都会执行以下步骤：

```
1. 启动 Office Workspace (127.0.0.1:3000)
2. 等待 Word Add-In 连接 (最多30秒)
   ⚠️  如果超时，请检查 Add-In 是否正确加载
3. 获取已连接文档列表
4. 执行替换动作
5. 验证返回结果
6. 清理资源（停止 Workspace）
```

---

## 快速开始

### 1. 运行单个测试

**从项目根目录运行（推荐）：**
```bash
# 运行文本替换测试 1
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test 1

# 运行格式替换测试 1
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test 1

# 运行边界情况测试 1
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test 1
```

**或者进入测试目录后运行：**
```bash
# 进入测试目录
cd manual_tests/word/replace_selection_e2e

# 运行文本替换测试 1
uv run python test_text_replace.py --test 1

# 运行格式替换测试 1
uv run python test_format_replace.py --test 1

# 运行边界情况测试 1
uv run python test_edge_cases.py --test 1
```

### 2. 运行所有测试

**从项目根目录运行（推荐）：**
```bash
# 运行所有文本替换测试（4个）
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test all

# 运行所有格式替换测试（4个）
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test all

# 运行所有边界情况测试（3个）
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/word/replace_selection_e2e

# 运行所有文本替换测试（4个）
uv run python test_text_replace.py --test all

# 运行所有格式替换测试（4个）
uv run python test_format_replace.py --test all

# 运行所有边界情况测试（3个）
uv run python test_edge_cases.py --test all
```

### 3. 完整测试流程

如果你想测试所有场景，依次运行：

**从项目根目录运行（推荐）：**
```bash
# 1. 文本替换
uv run python manual_tests/word/replace_selection_e2e/test_text_replace.py --test all

# 2. 格式替换
uv run python manual_tests/word/replace_selection_e2e/test_format_replace.py --test all

# 3. 边界情况
uv run python manual_tests/word/replace_selection_e2e/test_edge_cases.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/word/replace_selection_e2e

# 1. 文本替换
uv run python test_text_replace.py --test all

# 2. 格式替换
uv run python test_format_replace.py --test all

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
📝 替换内容: '...'
✅ 替换成功
   返回数据: {'replaced': true, 'characterCount': 11}
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

3. **选择为空**
   ```
   ❌ 替换失败: SELECTION_EMPTY
   错误码: 3002
   ```
   解决方法：在 Word 中选中一些文本后再运行测试

4. **替换失败**
   ```
   ❌ 替换失败: {error details}
   ```
   解决方法：检查 Word 文档是否可编辑，是否只读模式

---

## 测试场景说明

### 文本替换测试

**目的：** 验证基本的文本替换功能是否正常工作。

**测试内容：**
- 简单短文本
- 多行文本
- 特殊字符（中文、英文、数字、符号）
- 长文本（性能测试）

**预期结果：** 选中的文本被正确替换为指定内容。

**准备工作：**
1. 在 Word 文档中输入一些文本
2. 选中需要替换的文本
3. 运行测试

---

### 格式替换测试

**目的：** 验证 `format` 参数是否正确应用。

**测试内容：**
- 单一格式测试（bold, italic, fontSize, fontName, color, underline）
- 组合格式测试

**预期结果：** 替换后的文本应该显示为指定的格式。

**验证方法：** 在 Word 中查看替换后的文本格式。

**准备工作：**
1. 在 Word 文档中输入一些文本
2. 选中需要替换的文本
3. 运行测试
4. 检查替换后的文本格式

---

### 边界情况测试

**目的：** 验证异常情况的处理。

**测试内容：**
- 空选择替换（预期失败）
- 空字符串替换（删除内容）
- 图片替换

**预期结果：**
- 测试 1：返回错误码 3002
- 测试 2：删除选中内容
- 测试 3：替换为图片

**准备工作：**
1. 测试 1：确保 Word 中没有选中任何文本
2. 测试 2：选中一些文本
3. 测试 3：选中一些文本

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

### Q3: 替换内容不对

**问题：** 替换后的内容与预期不符

**解决方法：**
1. 确认已选中正确的文本
2. 检查 content 参数是否正确
3. 重新运行测试

### Q4: 格式未应用

**问题：** 替换后的文本没有显示格式

**解决方法：**
1. 检查 Add-In 是否支持格式设置
2. 确认 format 参数格式正确（JSON 对象）
3. 在 Word 中手动检查文本格式设置

### Q5: 空选择测试失败

**问题：** 测试 1（空选择）应该失败但实际成功了

**解决方法：**
1. 确认 Word 中没有选中任何文本
2. 点击文档空白处取消选择
3. 重新运行测试

---

## 调试技巧

### 1. 查看详细日志

测试会输出详细的执行日志，包括：
- Workspace 启动状态
- Add-In 连接状态
- 文档列表
- 执行的动作
- 返回结果

### 2. 手动验证

测试完成后，在 Word 中手动检查：
- 选中的文本是否被替换
- 替换后的内容是否正确
- 格式是否应用

### 3. 截图保存

建议在测试前后截图，方便对比结果。

### 4. 对比测试

运行多个测试后，对比不同参数的效果：
```bash
# 运行所有文本替换测试
uv run python test_text_replace.py --test all

# 检查 Word 文档中的所有替换结果
```

---

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/pages/30605313) - word:replace:selection 事件规范
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py) - Word DTOs
- [Namespace 实现](../../office4ai/environment/workspace/socketio/namespaces/word.py) - Word 命名空间
- [单元测试](../../tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py) - DTO 单元测试
- [Socket.IO API 标准](https://turingfocus.atlassian.net/wiki/spaces/OFFICE4AI) - 完整协议规范

---

## 最后更新

2026-01-09
