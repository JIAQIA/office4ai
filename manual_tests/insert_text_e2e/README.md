# Word Insert Text End-to-End Tests

测试 `word:insert:text` 事件的各种入参组合。

## 目录结构

```
insert_text_e2e/
├── __init__.py                 # 包初始化文件
├── test_basic_insert.py        # 基本插入测试（4个测试）
├── test_location_insert.py     # 位置插入测试（4个测试）
├── test_format_insert.py       # 格式插入测试（6个测试）
└── README.md                   # 本文档
```

## 测试概览

### 1. 基本插入测试 (`test_basic_insert.py`)

测试基础的文本插入功能，使用默认参数。

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 简单文本插入 | 插入 "Hello World" 到光标位置 |
| 2 | 多行文本插入 | 插入包含换行符的文本 |
| 3 | 特殊字符插入 | 插入特殊字符 `@#$%^&*()_+-=[]{}|;':",./<>?~\`` |
| 4 | 长文本插入 | 插入较长的文本段落（约200字符） |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test 1
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test 2
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test 3
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test 4

# 运行所有测试
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/insert_text_e2e

# 运行单个测试
uv run python test_basic_insert.py --test 1
uv run python test_basic_insert.py --test 2
uv run python test_basic_insert.py --test 3
uv run python test_basic_insert.py --test 4

# 运行所有测试
uv run python test_basic_insert.py --test all
# 如果在项目根目录
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test all
```

---

### 2. 位置插入测试 (`test_location_insert.py`)

测试不同插入位置参数的效果。

| 测试编号 | 测试名称 | location 参数 | 描述 |
|---------|---------|--------------|------|
| 1 | 光标位置插入 | `Cursor` | 在光标位置插入文本（默认值） |
| 2 | 文档开头插入 | `Start` | 在文档最开头插入文本 |
| 3 | 文档末尾插入 | `End` | 在文档最末尾插入文本 |
| 4 | 连续多次插入 | 混合 | 连续插入3次，每次使用不同位置 |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test 1
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test 2
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test 3
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test 4

# 运行所有测试
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/insert_text_e2e

# 运行单个测试
uv run python test_location_insert.py --test 1
uv run python test_location_insert.py --test 2
uv run python test_location_insert.py --test 3
uv run python test_location_insert.py --test 4

# 运行所有测试
uv run python test_location_insert.py --test all
```

---

### 3. 格式插入测试 (`test_format_insert.py`)

测试带格式参数的文本插入功能。

| 测试编号 | 测试名称 | format 参数 | 描述 |
|---------|---------|------------|------|
| 1 | 粗体文本 | `{bold: true}` | 插入粗体文本 |
| 2 | 斜体文本 | `{italic: true}` | 插入斜体文本 |
| 3 | 字体大小 | `{fontSize: 12/16/24}` | 插入3个不同字体大小的文本 |
| 4 | 字体名称 | `{fontName: "Arial"/"Times New Roman"/"Courier New"}` | 插入3个不同字体的文本 |
| 5 | 颜色设置 | `{color: "#FF0000"/"#00FF00"/"#0000FF"}` | 插入3个不同颜色的文本 |
| 6 | 组合格式 | 完整格式 | 插入带所有格式的文本（粗体+斜体+大小+字体+颜色） |

**运行方式：**

**方式一：从项目根目录运行（推荐）**
```bash
# 运行单个测试
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 1
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 2
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 3
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 4
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 5
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 6

# 运行所有测试
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test all
```

**方式二：进入测试目录后运行**
```bash
# 进入测试目录
cd manual_tests/insert_text_e2e

# 运行单个测试
uv run python test_format_insert.py --test 1
uv run python test_format_insert.py --test 2
uv run python test_format_insert.py --test 3
uv run python test_format_insert.py --test 4
uv run python test_format_insert.py --test 5
uv run python test_format_insert.py --test 6

# 运行所有测试
uv run python test_format_insert.py --test all
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
   - 对于光标位置插入测试，请将光标放在合适的位置
   - 建议使用空白文档进行测试

---

## 测试流程

每个测试都会执行以下步骤：

```
1. 启动 Office Workspace (127.0.0.1:3000)
2. 等待 Word Add-In 连接 (最多30秒)
   ⚠️  如果超时，请检查 Add-In 是否正确加载
3. 获取已连接文档列表
4. 执行插入动作
5. 验证返回结果
6. 清理资源（停止 Workspace）
```

---

## 快速开始

### 1. 运行单个测试

**从项目根目录运行（推荐）：**
```bash
# 运行基本插入测试 1
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test 1

# 运行位置插入测试 2
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test 2

# 运行格式插入测试 6（组合格式）
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 6
```

**或者进入测试目录后运行：**
```bash
# 进入测试目录
cd manual_tests/insert_text_e2e

# 运行基本插入测试 1
uv run python test_basic_insert.py --test 1

# 运行位置插入测试 2
uv run python test_location_insert.py --test 2

# 运行格式插入测试 6（组合格式）
uv run python test_format_insert.py --test 6
```

### 2. 运行所有测试

**从项目根目录运行（推荐）：**
```bash
# 运行所有基本插入测试（4个）
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test all

# 运行所有位置插入测试（4个）
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test all

# 运行所有格式插入测试（6个）
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/insert_text_e2e

# 运行所有基本插入测试（4个）
uv run python test_basic_insert.py --test all

# 运行所有位置插入测试（4个）
uv run python test_location_insert.py --test all

# 运行所有格式插入测试（6个）
uv run python test_format_insert.py --test all
```

### 3. 完整测试流程

如果你想测试所有场景，依次运行：

**从项目根目录运行（推荐）：**
```bash
# 1. 基本插入
uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test all

# 2. 位置插入
uv run python manual_tests/insert_text_e2e/test_location_insert.py --test all

# 3. 格式插入
uv run python manual_tests/insert_text_e2e/test_format_insert.py --test all
```

**或者进入测试目录后运行：**
```bash
cd manual_tests/insert_text_e2e

# 1. 基本插入
uv run python test_basic_insert.py --test all

# 2. 位置插入
uv run python test_location_insert.py --test all

# 3. 格式插入
uv run python test_format_insert.py --test all
```

---

## 预期结果

### ✅ 成功情况

```
✅ Workspace 启动成功
✅ Add-In 已连接
✅ 使用文档: file:///path/to/document.docx
📝 插入文本: '...'
✅ 插入成功
   返回数据: {'inserted': true}
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

3. **插入失败**
   ```
   ❌ 插入失败: {error details}
   ```
   解决方法：检查 Word 文档是否可编辑，是否只读模式

---

## 测试场景说明

### 基本插入测试

**目的：** 验证基本的文本插入功能是否正常工作。

**测试内容：**
- 简单短文本
- 多行文本
- 特殊字符
- 长文本（性能测试）

**预期结果：** 所有文本都能正确插入到文档中。

---

### 位置插入测试

**目的：** 验证 `location` 参数是否正确工作。

**测试内容：**
- `Cursor` - 在光标位置插入
- `Start` - 在文档开头插入
- `End` - 在文档末尾插入
- 连续插入 - 验证多个位置参数的累积效果

**预期结果：**
- `Cursor`: 文本插入在当前光标位置
- `Start`: 文本插入在文档最开始（原有内容之前）
- `End`: 文本插入在文档最后（原有内容之后）

**验证方法：** 在 Word 中查看插入位置是否正确。

---

### 格式插入测试

**目的：** 验证 `format` 参数是否正确应用。

**测试内容：**
- 单一格式测试（bold, italic, fontSize, fontName, color）
- 组合格式测试（所有格式同时应用）

**预期结果：** 插入的文本应该显示为指定的格式。

**验证方法：** 在 Word 中选中新插入的文本，检查其格式设置。

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

### Q3: 插入位置不对

**问题：** 文本没有插入到预期位置

**解决方法：**
1. 对于 `Cursor` 位置，请先将光标移动到目标位置
2. 等待测试开始后再移动光标（测试有5秒等待时间）
3. 检查 location 参数是否正确

### Q4: 格式未应用

**问题：** 插入的文本没有显示格式

**解决方法：**
1. 检查 Add-In 是否支持格式设置
2. 确认 format 参数格式正确（JSON 对象）
3. 在 Word 中手动检查文本格式设置

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
- 文本是否插入
- 插入位置是否正确
- 格式是否应用

### 3. 截图保存

建议在测试前后截图，方便对比结果。

---

## 相关文档

- [Confluence 文档](https://turingfocus.atlassian.net/wiki/x/DADGAQ)
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py)
- [Namespace 实现](../../office4ai/environment/workspace/socketio/namespaces/word.py)
- [单元测试](../../tests/unit_tests/office4ai/environment/workspace/dtos/test_word_dtos.py)

---

## 最后更新

2026-01-09
