# Word Get Document Stats End-to-End Tests (自动化版本)

测试 `word:get:documentStats` 事件的各种场景。

## 特性

- **自动复制测试文档** - 每次测试使用独立副本，不影响原始夹具
- **自动打开 Word** - 通过系统命令打开文档
- **自动验证结果** - 预定义期望值和自定义验证器
- **智能清理** - 成功后自动删除副本，失败则保留供调试

## 目录结构

```
get_document_stats_e2e/
├── test_basic_stats.py           # 测试脚本
├── README.md                     # 本文档
└── ../fixtures/get_document_stats_e2e/  # 测试夹具
    ├── empty.docx                # 空白文档
    ├── simple.docx               # 简单文本文档
    ├── complex.docx              # 复杂文档（表格、列表）
    └── large.docx                # 大文档（~10页）
```

## 快速开始

### 前置条件

1. **Word Add-In 已安装** - 确保 `office-editor4ai` Add-In 已在 Word 中加载
2. **Add-In 配置正确** - 连接地址为 `http://127.0.0.1:3000`

### 运行单个测试

```bash
# 测试 1: 空白文档
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --test 1

# 测试 2: 简单文档
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --test 2

# 测试 3: 复杂文档
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --test 3

# 测试 4: 大文档
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --test 4
```

### 运行所有测试

```bash
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --test all
```

### 列出所有测试

```bash
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --list
```

## 命令行选项

| 选项 | 说明 |
|------|------|
| `--test {1,2,3,4,all}` | 选择要运行的测试 |
| `--no-auto-open` | 不自动打开文档（手动打开模式） |
| `--always-cleanup` | 无论成功失败都清理测试文件 |
| `--list` | 列出所有测试用例 |

## 测试用例

| 编号 | 名称 | 夹具 | 验证内容 |
|------|------|------|----------|
| 1 | 空白文档统计 | `empty.docx` | wordCount=0, characterCount=0, paragraphCount≈1 |
| 2 | 简单文档统计 | `simple.docx` | 所有统计值 > 0 |
| 3 | 复杂文档统计 | `complex.docx` | 所有统计值 > 0 |
| 4 | 大文档统计 | `large.docx` | wordCount > 1000, characterCount > 5000 |

## 测试流程

```
1. 复制夹具文件到临时目录
2. 通过系统命令打开文档（macOS: open, Windows: start）
3. 等待 Word 加载（~3秒）
4. 启动 Workspace 服务器
5. 等待 Add-In 连接（最多 30 秒）
6. 执行获取文档统计动作
7. 验证返回结果
8. 成功则清理临时文件，失败则保留
```

## 手动模式

如果自动打开不工作，可以使用手动模式：

```bash
uv run python manual_tests/word/get_document_stats_e2e/test_basic_stats.py --test 1 --no-auto-open
```

在此模式下，你需要：
1. 手动打开测试脚本打印的文件路径
2. 等待 Add-In 连接
3. 测试会自动继续执行

## 调试失败的测试

当测试失败时，临时文件会被保留在 `/tmp/office4ai_e2e_tests/` 目录：

```bash
# 查看保留的测试文件
ls -la /tmp/office4ai_e2e_tests/

# 手动清理
uv run python manual_tests/e2e_base.py --clean-temp
```

## 自定义测试夹具

如果需要修改测试夹具，可以：

1. 直接编辑 `manual_tests/word/fixtures/get_document_stats_e2e/` 下的 `.docx` 文件
2. 或重新生成夹具：

```bash
# 删除现有夹具
rm -rf manual_tests/word/fixtures/get_document_stats_e2e/*.docx

# 重新生成
uv run python manual_tests/e2e_base.py --create-fixtures manual_tests/word/fixtures/get_document_stats_e2e
```

## 预期输出示例

```
======================================================================
🧪 测试 2: 简单文档统计
======================================================================
📋 描述: 简单文本文档应该正确统计字数、字符数和段落数
📄 夹具: simple.docx
📄 创建工作副本: simple_20260203_143052.docx
📂 打开文档: simple_20260203_143052.docx
✅ Workspace 启动成功
⏳ 等待 Word Add-In 连接...
✅ Add-In 已连接

📝 执行: 获取文档统计...

⏱️  执行时间: 125.3ms
✅ 获取成功
   字数: 42
   字符数: 186
   段落数: 4

📊 验证结果:
   ✅ 自定义验证通过

======================================================================
✅ 测试 2 通过
======================================================================
🧹 已清理: simple_20260203_143052.docx
```

## 常见问题

### Q1: Add-In 连接超时

**问题**: `超时：未检测到 Add-In 连接`

**解决方案**:
1. 确认 Word Add-In 已正确加载
2. 检查 Add-In 控制台是否有错误
3. 确认连接地址为 `http://127.0.0.1:3000`

### Q2: 文档无法自动打开

**问题**: macOS 上文档无法自动打开

**解决方案**:
1. 确认 Microsoft Word 已安装并设为 `.docx` 默认应用
2. 使用 `--no-auto-open` 选项手动打开

### Q3: 测试文件未清理

**问题**: 临时文件越来越多

**解决方案**:
```bash
# 手动清理临时文件
uv run python manual_tests/e2e_base.py --clean-temp
```

## 相关文档

- [E2E 测试基础设施](../e2e_base.py) - 测试运行器和夹具管理
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py) - 数据结构
- [Namespace 实现](../../office4ai/environment/workspace/socketio/namespaces/word.py) - 事件处理

---

**最后更新**: 2026-02-03
