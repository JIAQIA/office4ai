# Word Get Document Structure End-to-End Tests

测试 `word:get:documentStructure` 事件的各种场景。

## 目录结构

```
get_document_structure_e2e/
├── __init__.py                 # 包初始化文件
├── test_basic_structure.py     # 基本结构获取测试（4个测试）
└── README.md                   # 本文档
```

## 特性

- **自动复制测试文档** - 使用临时副本，不影响原始夹具
- **自动打开 Word** - 无需手动打开文档
- **自动验证结果** - 比对预期值和实际返回
- **自动清理** - 成功后删除临时文件，失败则保留供调试
- **正确的清理顺序** - 先关闭文档再停止 Workspace，避免 2 分钟等待

## 测试概览

| 测试编号 | 测试名称 | 描述 | 夹具 |
|---------|---------|------|------|
| 1 | 空白文档结构 | 获取空白文档的结构信息 | empty.docx |
| 2 | 简单文档结构 | 获取包含简单文本的文档结构 | simple.docx |
| 3 | 复杂文档结构 | 获取包含表格和列表的文档结构 | complex.docx |
| 4 | 大文档结构 | 获取包含大量段落的文档结构（性能测试） | large.docx |

## 快速开始

### 运行单个测试

```bash
# 运行测试 1（空白文档）
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test 1

# 运行测试 2（简单文档）
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test 2

# 运行测试 3（复杂文档）
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test 3

# 运行测试 4（大文档）
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test 4
```

### 运行所有测试

```bash
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test all
```

### 列出所有测试

```bash
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --list
```

### 其他选项

```bash
# 手动打开文档模式（不自动打开）
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test 1 --no-auto-open

# 失败时也清理文件
uv run python manual_tests/word/get_document_structure_e2e/test_basic_structure.py --test 1 --always-cleanup
```

## 前置条件

1. **Word Add-In 已加载** - 在 Word 中加载 `office-editor4ai` Add-In
2. **python-docx 已安装** - 用于创建测试夹具

## 预期输出

### 成功示例

```
======================================================================
🧪 测试 1: 空白文档结构
======================================================================
📋 描述: 空白文档应该有 0-1 段落（Word 默认有一个空段落），0 表格，0 图片，1 节
📄 夹具: empty.docx
📄 创建工作副本: empty_20260203_160000.docx
📂 打开文档: empty_20260203_160000.docx
✅ Workspace 启动成功
⏳ 等待 Word Add-In 连接...
✅ Add-In 已连接

📝 执行: 获取文档结构...

⏱️  执行时间: 350.6ms
✅ 获取成功
   段落数: 1
   表格数: 0
   图片数: 0
   节数: 1

📊 验证结果:
   ✅ 段落数: 1 (预期 1)
   ✅ 表格数: 0 (预期 0)
   ✅ 图片数: 0 (预期 0)
   ✅ 节数: 1 (预期 1)

======================================================================
✅ 测试 1 通过
======================================================================
📕 已关闭文档: empty_20260203_160000.docx
🧹 已清理: empty_20260203_160000.docx
```

## 测试夹具

测试夹具自动创建在 `manual_tests/word/fixtures/get_document_structure_e2e/` 目录：

| 文件 | 内容 |
|------|------|
| empty.docx | 空白文档 |
| simple.docx | 包含标题和 3 个段落的简单文档 |
| complex.docx | 包含标题、段落、表格和列表的复杂文档 |
| large.docx | 包含约 30 个段落的大文档 |

工作副本存放在 `manual_tests/.test_working/` 目录（已加入 .gitignore）。

## 常见问题

### Q1: Add-In 连接超时

**问题：** `超时：未检测到 Add-In 连接`

**解决方法：**
1. 确认 Word Add-In 已加载
2. 检查 Add-In 是否能访问 `http://127.0.0.1:3000`
3. 打开浏览器开发者工具查看控制台错误

### Q2: 测试后等待很长时间

**问题：** 测试完成后等待约 2 分钟才退出

**原因：** 清理顺序错误。如果先停止 Workspace 而文档还开着，需要等待 Socket.IO 连接超时。

**解决方法：** 本测试已修复此问题，会先关闭文档再停止 Workspace。

### Q3: 文档结构计数不准确

**问题：** 返回的计数与实际不符

**解决方法：**
1. 手动检查文档中的元素数量
2. Word 的段落计数可能包含空段落
3. 图片计数可能包含不同类型的图片（嵌入式、浮动等）

## 相关文件

- [E2E 基础设施](../e2e_base.py) - 测试运行器和夹具管理
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py) - 数据结构定义
- [连接调试 SKILL](../../.claude/skills/debug-socketio-connection/SKILL.md) - 连接问题排查指南

## 最后更新

2026-02-03
