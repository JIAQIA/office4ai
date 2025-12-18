# Office4AI

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Office4AI** 是一个专为 AI Agent 设计的强大 Office 文档管理环境，提供对 Word (docx)、Excel (xlsx)、PowerPoint (pptx) 文档的智能编辑和管理功能。

## ✨ 特性

- 📄 **Word 文档管理** - 创建、编辑、格式化 Word 文档（.docx）
- 📊 **Excel 表格处理** - 数据读写、公式计算、图表生成（.xlsx）
- 📽️ **PowerPoint 演示文稿** - 幻灯片创建、内容编辑、样式设置（.pptx）
- 🔧 **LibreOffice 集成** - 完整的 LibreOffice API 支持
- 🖥️ **终端环境** - 本地和 Docker 容器内的命令执行
- 📁 **工作区管理** - 文件系统操作、目录树浏览
- 🎯 **为 AI 优化** - 专门设计的接口，方便 AI Agent 理解和操作文档

## 🎯 设计目标

Office4AI 的核心设计理念是为 AI Agent 提供一个**高内聚、低耦合**的文档操作环境：

- **高内聚**：所有 Office 功能（编辑、格式化、转换）都集中在统一的接口中
- **低耦合**：独立于任何特定的 AI 框架，可以轻松集成到不同的 Agent 系统
- **Gymnasium 兼容**：实现了 Gymnasium Env 接口，可作为强化学习环境使用

## 📦 安装

### ⚠️ 系统依赖要求

**在安装 Office4AI 之前，请先安装 LibreOffice：**

<details>
<summary><b>📥 LibreOffice 安装指南（点击展开）</b></summary>

#### macOS
```bash
brew install --cask libreoffice
```

#### Ubuntu/Debian
```bash
sudo apt-get update
sudo apt-get install libreoffice libreoffice-script-provider-python
```

#### Fedora/RHEL
```bash
sudo dnf install libreoffice libreoffice-pyuno
```

#### Arch Linux
```bash
sudo pacman -S libreoffice-fresh
```

#### Windows
从 [LibreOffice 官网](https://www.libreoffice.org/download/download/) 下载并安装

</details>

### 使用 uv（推荐）

```bash
# 克隆仓库
git clone https://github.com/JQQ/office4ai.git
cd office4ai

# 安装依赖
uv sync

# 开发模式安装
uv sync --all-extras
```

### 使用 pip

```bash
pip install office4ai
```

## 🚀 快速开始

### 基础使用

```python
from office4ai import OfficeEnv, OfficeAction

# 创建 Office 环境实例
env = OfficeEnv(
    root_dir="/path/to/your/documents",
    project_name="my_documents",
    libreoffice_path="/Applications/LibreOffice.app/Contents/MacOS/soffice"
)

# 打开 Word 文档
action = OfficeAction(
    category="document",
    action_name="open_docx",
    action_args={"file_path": "document.docx"}
)
obs, reward, done, truncated, info = env.step(action.model_dump())

# 编辑文档内容
edit_action = OfficeAction(
    category="document",
    action_name="edit_docx",
    action_args={
        "file_path": "document.docx",
        "operations": [{
            "type": "add_paragraph",
            "text": "Hello, AI World!",
            "style": "Heading 1"
        }]
    }
)
obs, reward, done, truncated, info = env.step(edit_action.model_dump())

# 处理 Excel 表格
excel_action = OfficeAction(
    category="spreadsheet",
    action_name="edit_xlsx",
    action_args={
        "file_path": "data.xlsx",
        "operations": [{
            "type": "set_cell",
            "sheet": "Sheet1",
            "cell": "A1",
            "value": "Hello Excel"
        }]
    }
)
obs, reward, done, truncated, info = env.step(excel_action.model_dump())
```

## 📚 核心概念

### Office Actions

Office4AI 支持三类操作：

1. **Document Actions** - Word 文档操作
   - `open_docx` - 打开文档
   - `edit_docx` - 编辑文档
   - `save_docx` - 保存文档
   - `format_docx` - 格式化文档
   - `convert_docx` - 转换文档格式

2. **Spreadsheet Actions** - Excel 表格操作
   - `open_xlsx` - 打开表格
   - `edit_xlsx` - 编辑表格
   - `save_xlsx` - 保存表格
   - `calculate_xlsx` - 计算公式
   - `chart_xlsx` - 创建图表

3. **Presentation Actions** - PowerPoint 演示文稿操作
   - `open_pptx` - 打开演示文稿
   - `edit_pptx` - 编辑演示文稿
   - `save_pptx` - 保存演示文稿
   - `add_slide_pptx` - 添加幻灯片

### LibreOffice 集成

Office4AI 提供了两种 LibreOffice 集成方式：

#### 1. UNO Bridge（推荐）- 双层客户端-服务器架构

UNO Bridge 采用双层架构，完全解耦 Python 版本依赖：

```
┌─────────────────────────────────────────────────────────────┐
│  Python 3.11+ 客户端 (你的项目)                              │
│  ┌────────────────────────────────────────────────────┐    │
│  │  使用 xmlrpc.client 调用服务器                       │    │
│  │  无需安装 LibreOffice 或 uno 库                     │    │
│  └────────────────────────────────────────────────────┘    │
└──────────────────────┬──────────────────────────────────────┘
                       │ XML-RPC over HTTP (端口 2003)
┌──────────────────────▼──────────────────────────────────────┐
│  Python 3.8 服务器 (LibreOffice 自带)                        │
│  ┌────────────────────────────────────────────────────┐    │
│  │  处理 XML-RPC 请求并调用 UNO API                    │    │
│  └────────────────────────────────────────────────────┘    │
└──────────────────────┬──────────────────────────────────────┘
                       │ UNO Bridge (端口 2002)
┌──────────────────────▼──────────────────────────────────────┐
│  LibreOffice 进程 (无头模式)                                  │
└─────────────────────────────────────────────────────────────┘
```

**优势**：
- ✅ **完全解耦**：项目使用 Python 3.11+，LibreOffice 使用 Python 3.8
- ✅ **依赖隔离**：项目依赖（如 MCP SDK）与 LibreOffice 依赖完全分离
- ✅ **易于维护**：LibreOffice 升级不影响项目代码
- ✅ **灵活部署**：可以将 LibreOffice 服务部署在不同的机器上

**使用方法**：

```bash
# 1. 启动 UNO Bridge 服务器（使用 LibreOffice Python）
./scripts/start_uno_server.sh

# 2. 在你的项目中使用客户端（使用项目 Python 3.11+）
```

```python
from office4ai.uno_bridge.client import UnoClient

# 创建客户端
client = UnoClient(server="127.0.0.1", port=2003)

# 测试连接
if client.ping():
    print("连接成功！")

# 替换文档中的文本
result = client.replace_text(
    file_path="/path/to/document.docx",
    search_text="旧文本",
    replace_text="新文本"
)

if result["success"]:
    print(f"替换成功！替换了 {result['count']} 处")
```

#### 2. 直接 UNO API 调用

- **UNO API** - 完整的 LibreOffice UNO API 支持
- **文档转换** - 支持多种格式之间的转换
- **批处理** - 批量处理多个文档

**注意**：直接使用 UNO API 需要项目 Python 版本与 LibreOffice Python 版本一致

## 🛠️ 开发

### 环境设置

```bash
# 安装开发依赖
uv sync

# 或使用 poe 任务
poe install-dev
```

### 常用命令

项目使用 [poethepoet](https://github.com/nat-n/poethepoet) 管理开发任务：

```bash
# 代码检查
poe lint              # 运行 ruff 检查
poe lint-fix          # 自动修复 lint 问题
poe format            # 格式化代码
poe format-check      # 检查代码格式

# 类型检查
poe typecheck         # 运行 mypy 类型检查

# 测试
poe test              # 运行所有测试
poe test-unit         # 仅运行单元测试
poe test-integration  # 仅运行集成测试
poe test-cov          # 运行测试并生成覆盖率报告
poe test-verbose      # 详细模式运行测试

# 组合任务
poe check             # 运行所有检查（lint + format-check + typecheck）
poe fix               # 自动修复问题（lint-fix + format）
poe pre-commit        # 提交前检查（format + lint-fix + typecheck + test）

# 清理
poe clean             # 清理缓存和临时文件
poe clean-pyc         # 清理 Python 缓存
poe clean-cov         # 清理覆盖率报告
```

### 运行测试

```bash
# 运行所有测试
poe test

# 运行特定测试
pytest tests/test_docx.py -v

# 生成覆盖率报告
poe test-cov
```

### 代码规范

项目使用以下工具确保代码质量：

- **Ruff** - 快速的 Python linter 和 formatter
- **MyPy** - 静态类型检查
- **Pytest** - 测试框架

提交代码前请运行：

```bash
poe pre-commit
```

## 🏗️ 架构设计

```
office4ai/
├── base.py                 # Office 环境基类
├── schema.py              # 数据模型定义
├── exceptions.py          # 异常类
├── utils.py              # 工具函数
├── dtos/                 # 数据传输对象
│   ├── base_protocol.py
│   ├── commands.py
│   └── documents.py
├── environment/          # 环境实现
│   ├── terminal/        # 终端环境
│   │   ├── base.py
│   │   └── local_terminal_env.py
│   └── workspace/       # 工作区
│       ├── base.py
│       └── utils.py
├── office/              # Office 实现
│   ├── docx_handler.py  # Word 文档处理
│   ├── xlsx_handler.py  # Excel 表格处理
│   ├── pptx_handler.py  # PowerPoint 处理
│   ├── libreoffice.py   # LibreOffice 集成
│   └── mcp/            # MCP 服务器
│       └── server.py
├── uno_bridge/          # UNO Bridge 双层架构
│   ├── __init__.py
│   ├── server.py        # 服务器端（Python 3.8）
│   └── client.py        # 客户端（Python 3.11+）
└── py.typed
```

## 🔌 MCP 服务器

Office4AI 提供了 MCP (Model Context Protocol) 服务器，可以轻松集成到支持 MCP 的 AI 应用中：

```bash
# 启动 MCP 服务器
office4ai-mcp --root-dir /path/to/documents
```

## 📖 文档

- [API 文档](docs/api.md)（待完善）
- [架构设计](docs/architecture.md)（待完善）
- [LibreOffice 集成指南](docs/libreoffice.md)（待完善）

## 🤝 贡献

欢迎贡献！请查看 [CONTRIBUTING.md](CONTRIBUTING.md)（待创建）了解详情。

### 贡献流程

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

- 基于 [Gymnasium](https://gymnasium.farama.org/) 环境接口
- LibreOffice UNO API 支持
- 灵感来源于 [IDE4AI](https://github.com/JQQ/ide4ai)

## 📮 联系方式

- 作者：JQQ
- Email：jqq1716@gmail.com
- GitHub：[@JQQ](https://github.com/JQQ)

## 🗺️ 路线图

- [ ] 完善文档和示例
- [ ] 支持更多文档格式（PDF、ODF 等）
- [ ] 添加文档模板系统
- [ ] 提供 Web UI 界面
- [ ] 性能优化和大型文档支持
- [ ] 更多 AI 框架集成示例
- [ ] 文档智能分析功能

---

**如果这个项目对你有帮助，请给个 ⭐️ Star！**
