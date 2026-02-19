"""
Office4AI 手动调试测试套件

=====================================
⚠️  重要提示：手动调试测试 ⚠️
=====================================

本目录包含需要手动配合的集成测试用例，用于调试和验证 Workspace 与 Office Add-In 的端到端通信。

这些测试用例具有以下特点：
- ✅ 需要手动启动外部服务（如 Word Add-In）
- ✅ 需要手动操作 Office 软件
- ❌ 无法完全自动化，不参与 CI/CD
- ❌ 不被 pytest 自动发现和执行

=====================================
目录结构
=====================================

manual_tests/                   # 项目根目录下的手动调试测试目录
├── __init__.py                 # 本文件：使用说明
├── e2e_base.py                # 共享 E2E 测试基础设施
├── test_workspace_startup.py  # Workspace 启动与连接测试
├── MANUAL_TEST.md             # 详细的手动测试指南
├── fixtures/                  # 测试文档 fixture 生成器
├── word/                      # Word E2E 测试
│   ├── test_helpers.py        # Word 测试辅助函数
│   ├── test_word_e2e.py       # Word 端到端集成测试
│   └── *_e2e/                 # 各功能 E2E 测试目录 (11 个)
└── ppt/                       # PPT E2E 测试
    ├── e2e_base.py            # PPT 测试基础设施
    ├── test_helpers.py        # PPT 测试辅助函数
    └── *_e2e/                 # 各功能 E2E 测试目录 (15 个)

=====================================
快速开始
=====================================

1. 前置条件
   ----------
   - Python 3.10+ (推荐 3.11)
   - 已安装项目依赖: `poe install-dev`
   - 已启动 Word Add-In（参考 MANUAL_TEST.md）

2. 运行测试
   ----------
   ⚠️  这些测试不会被 pytest 自动发现，必须手动运行：

   # 方式 1: 直接使用 uv run（推荐）
   uv run python manual_tests/test_workspace_startup.py

   # 方式 2: 使用 python -m
   python -m manual_tests.test_workspace_startup

   # 方式 3: 切换到目录后运行
   cd manual_tests
   uv run python test_workspace_startup.py

3. 带参数运行
   ----------
   # Word 端到端测试（支持 --mode 参数）
   uv run python manual_tests/word/test_word_e2e.py --mode health   # 快速健康检查
   uv run python manual_tests/word/test_word_e2e.py --mode e2e     # 完整端到端测试

=====================================
测试用例说明
=====================================

1. test_workspace_startup.py
   --------------------------
   用途: 测试 Workspace 启动和 Add-In 连接

   启动命令:
       uv run python manual_tests/test_workspace_startup.py

   测试流程:
       1. 启动 Workspace Socket.IO 服务器 (http://127.0.0.1:3000)
       2. 等待 Word Add-In 连接（最长 5 分钟）
       3. 显示已连接文档列表
       4. 保持运行直到手动停止（Ctrl+C 或 Enter）

   适用场景:
       - 调试 Add-In 连接问题
       - 验证 Socket.IO 通信
       - 手动测试 Add-In 功能

2. test_word_e2e.py
   --------------------------
   用途: 完整的 Word 端到端集成测试

   启动命令:
       # 健康检查模式（快速验证 Workspace 能否启动）
       uv run python manual_tests/word/test_word_e2e.py --mode health

       # 端到端测试模式（完整测试请求-响应流程）
       uv run python manual_tests/word/test_word_e2e.py --mode e2e

   测试流程 (e2e 模式):
       1. 启动 Workspace
       2. 等待 Add-In 连接（30 秒超时）
       3. 获取已连接文档列表
       4. 调用 word:get:selectedContent 获取选中内容
       5. 验证返回结果
       6. 清理资源

   适用场景:
       - 验证完整的请求-响应流程
       - 测试特定 Word 功能
       - 回归测试

3. MANUAL_TEST.md
   --------------------------
   用途: 详细的手动测试指南

   查看:
       cat manual_tests/MANUAL_TEST.md
       # 或在文本编辑器中打开

   内容:
       - 测试环境要求
       - 测试准备步骤
       - 测试场景详解
       - 常见问题排查
       - 测试报告模板

=====================================
CI/CD 配置
=====================================

pytest 配置（pyproject.toml）应排除本目录：

    [tool.pytest.ini_options]
    testpaths = ["tests"]  # 会自动发现所有 test_*.py
    # 手动测试不会被 pytest 执行，因为：
    # 1. 可以通过 conftest.py 排除
    # 2. 或者不在 testpaths 中包含 manual_tests

推荐配置方式（二选一）：

方式 1: 在 tests/conftest.py 中添加收集排除
    ```python
    def pytest_collection_modifyitems(config, items):
        # 排除 manual_tests 目录
        items[:] = [item for item in items
                    if "/manual_tests/" not in item.fspath.strpath]
    ```

方式 2: 修改 pytest.ini_options 的 testpaths
    ```toml
    [tool.pytest.ini_options]
    testpaths = ["tests/unit_tests", "tests/integration_tests"]
    ```

=====================================
常见问题
=====================================

Q: 为什么这些测试不被 pytest 自动发现？
A: 因为它们需要手动配合启动外部服务和 Office 软件，无法在 CI/CD 中自动运行。

Q: 如何确认这些测试不会在 CI/CD 中运行？
A: 检查 pytest 配置（pyproject.toml）和 tests/conftest.py，确保排除 manual_tests 目录。

Q: 可以将自动化的测试放在这里吗？
A: 不建议。请将自动化测试放在：
    - tests/unit_tests/      # 单元测试（独立功能）
    - tests/integration_tests/  # 集成测试（可自动化）

Q: 如何添加新的手动调试测试？
A: 在项目根目录的 manual_tests/ 下创建 test_*.py 文件，
   并在本文件顶部的"目录结构"部分添加说明。

=====================================
相关文档
=====================================

- manual_tests/MANUAL_TEST.md - 详细的手动测试指南
- docs/workspace_implementation_plan.md - Workspace 实现计划
- docs/mvp_implementation_plan.md - MVP 实现计划
- CLAUDE.md - 项目元信息

=====================================
更新日志
=====================================

2026-01-08 - 创建手动调试测试目录（项目根目录）
    - 从 tests/integration_tests/ 移动需要手动配合的测试
    - 添加本使用说明文件
    - 添加到 .gitignore（本地调试用，不提交到版本控制）
    - 确保手动测试不参与 CI/CD

=====================================
"""

__all__ = []  # 不导出任何内容，本目录仅用于组织测试文件
