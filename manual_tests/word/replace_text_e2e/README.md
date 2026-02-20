# Word Replace Text End-to-End Tests

测试 `word:replace:text` 事件的各种入参组合。

## 目录结构

```
replace_text_e2e/
├── __init__.py                 # 包初始化文件
├── test_basic_replace.py       # 基础替换测试（6个测试）
├── test_format_replace.py      # 格式化替换测试（6个测试）
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
| 4 | 多行文本替换 | "line1\nline2" | "new\ncontent" | 替换多行文本 |
| 5 | 特殊字符替换 | "Café" | "Coffee" | 替换含特殊字符的文本 |
| 6 | 长文本替换 | 长文本 | 长文本 | 替换较长段落 |

**运行方式：**

```bash
# 运行单个测试
uv run python manual_tests/word/replace_text_e2e/test_basic_replace.py --test 1

# 运行所有测试
uv run python manual_tests/word/replace_text_e2e/test_basic_replace.py --test all

# 列出所有测试
uv run python manual_tests/word/replace_text_e2e/test_basic_replace.py --list
```

---

### 2. 格式化替换测试 (`test_format_replace.py`)

测试 `format` 参数的格式化替换功能，使用相同文本替换 + format 参数实现"为已有文本添加格式"。

| 测试编号 | 测试名称 | format | 描述 |
|---------|---------|--------|------|
| 1 | 粗体格式化 | `{bold: true}` | 搜索 'important' 替换为自身 + 粗体 |
| 2 | 斜体格式化 | `{italic: true}` | 搜索 'emphasis' 替换为自身 + 斜体 |
| 3 | 颜色格式化 | `{color: "#FF0000"}` | 搜索 'alert' 替换为自身 + 红色 |
| 4 | styleName 样式格式化 | `{styleName: "Heading 2"}` | 搜索 'Chapter' 替换为自身 + 标题2样式 |
| 5 | 组合格式化 | `{bold, italic, color, fontSize}` | 搜索 'Critical' + 多种格式同时应用 |
| 6 | 替换文本 + 格式同时应用 | `{bold, color}` | 搜索 'alert' 替换为 'WARNING' + 粗体红色 |

**运行方式：**

```bash
# 运行单个测试
uv run python manual_tests/word/replace_text_e2e/test_format_replace.py --test 1

# 运行所有测试
uv run python manual_tests/word/replace_text_e2e/test_format_replace.py --test all

# 列出所有测试
uv run python manual_tests/word/replace_text_e2e/test_format_replace.py --list
```

---

## 前置条件

1. **Word Add-In 已加载** - 在 Word 中加载 `office-editor4ai` Add-In
2. **python-docx 已安装** - 用于创建测试夹具和验证文档内容

## 测试夹具

测试夹具自动创建在 `manual_tests/word/fixtures/replace_text_e2e/` 目录：

| 文件 | 用途 |
|------|------|
| replace_targets.docx | 基础替换测试用（含 "old"×5, "test"×5, "delete"×3, "Café"×3, 长段落×2） |
| format_targets.docx | 格式化替换测试用（含 "important"×3, "emphasis"×3, "alert"×3, "Chapter"×2, "Critical"×2） |

工作副本存放在 `manual_tests/.test_working/` 目录（已加入 .gitignore）。

## 命令行选项

| 选项 | 说明 |
|------|------|
| `--test {1..6,all}` | 选择要运行的测试 |
| `--no-auto-open` | 不自动打开文档（手动打开模式） |
| `--always-cleanup` | 无论成功失败都清理测试文件 |
| `--list` | 列出所有测试用例 |

## 相关文档

- [E2E 基础设施](../e2e_base.py) - 测试运行器和夹具管理
- [DTO 定义](../../office4ai/environment/workspace/dtos/word.py) - Word DTOs
- [Namespace 实现](../../office4ai/environment/workspace/socketio/namespaces/word.py) - Word 命名空间

---

**最后更新**: 2026-02-20
