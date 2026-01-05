# CLAUDE.md - Office4AI 项目元信息

> 本文档仅为 AI 助手提供项目元信息索引。详细内容请参考相应文件。

---

## 项目定位

**Office4AI** - 专为 AI Agent 设计的 Office 文档管理环境
- **协议**: MCP (Model Context Protocol) Server
- **接口**: Gymnasium 环境接口
- **架构**: 对齐 `examples/ide4ai` 分层架构
- **集成**: LibreOffice UNO Bridge 双层架构

---

## 工具链

| 类别 | 工具 | 配置位置 |
|------|------|---------|
| 依赖管理 | uv | `uv.lock` |
| 任务运行 | poethepoet | `pyproject.toml: [tool.poe.tasks]` |
| Lint/Format | ruff | `pyproject.toml: [tool.ruff]` |
| 类型检查 | mypy | `pyproject.toml: [tool.mypy]` |
| 测试框架 | pytest | `pyproject.toml: [tool.pytest.ini_options]` |

---

## 核心命令速查

```bash
# 安装
poe install-dev          # uv sync --all-extras

# 代码质量
poe lint                 # ruff check
poe lint-fix             # ruff check --fix
poe format               # ruff format
poe typecheck            # mypy office4ai

# 测试
poe test                 # 全部测试
poe test-unit            # 单元测试
poe test-integration     # 集成测试 (-m integration)
poe test-cov             # 覆盖率报告

# 组合
poe check                # lint + format-check + typecheck
poe fix                  # lint-fix + format
poe pre-commit           # format + lint-fix + typecheck + test

# 清理
poe clean                # clean-pyc + clean-cov
```

---

## 项目结构

```
office4ai/
├── base.py              # OfficeEnv 基类
├── schema.py            # 数据模型 (OfficeAction, OfficeObservation)
├── exceptions.py        # 异常定义
├── utils.py             # 工具函数
├── dtos/                # 数据传输对象
├── environment/         # TerminalEnv, WorkspaceEnv
├── office/              # Office 处理器 (docx/xlsx/pptx)
├── a2c_smcp/            # MCP Server 基础设施
│   └── server.py        # BaseMCPServer
└── uno_bridge/          # UNO Bridge 双层架构
    ├── server.py        # XML-RPC Server (LibreOffice Python 3.8)
    └── client.py        # XML-RPC Client (项目 Python 3.11+)
```

---

## 关键约定

### 代码规范
- **行长度**: 120 字符
- **Python 版本**: 3.10+ (推荐 3.11)
- **类型注解**: 强制 (mypy `disallow_untyped_defs = true`)
- **导入顺序**: 标准库 → 第三方 → 本地 (ruff 自动)

### 测试组织
```
tests/
├── conftest.py          # 公共 fixtures
├── unit_tests/          # 单元测试 (独立功能)
└── integration_tests/   # 集成测试 (@pytest.mark.integration)
```

### Ruff 规则
```toml
select = ["E", "W", "F", "I", "B", "C4", "UP"]
ignore = ["E501", "B008", "C901"]
"__init__.py" = ["F401"]  # 允许未使用导入
```

---

## 参考文档

| 文档 | 用途 |
|------|------|
| [README.md](README.md) | 项目介绍、快速开始 |
| [pyproject.toml](pyproject.toml) | 完整配置 |
| [docs/office4ai_dev_plan.md](docs/office4ai_dev_plan.md) | 开发方案 |
| [docs/ide4ai_alignment_plan.md](docs/ide4ai_alignment_plan.md) | A2C 协议对齐 |

---

## 外部资源

- [uv](https://github.com/astral-sh/uv)
- [Poe the Poet](https://github.com/nat-n/poethepoet)
- [Ruff](https://docs.astral.sh/ruff/)
- [Gymnasium](https://gymnasium.farama.org/)
- [MCP Protocol](https://modelcontextprotocol.io)
- [examples/ide4ai](examples/ide4ai) - 架构参考

---

**最后更新**: 2026-01-03
**维护者**: JQQ <jqq1716@gmail.com>
