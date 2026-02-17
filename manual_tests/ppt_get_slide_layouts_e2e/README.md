# PPT Get Slide Layouts E2E Tests

`ppt:get:slideLayouts` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 默认布局列表 | simple.pptx | 获取所有可用布局，验证列表不为空 |
| 2 | 布局名称验证 | simple.pptx | 验证每个布局都包含 name 字段 |
| 3 | 包含占位符选项 | simple.pptx | 带 includePlaceholders=true 获取布局详情 |

## 运行

```bash
uv run python manual_tests/ppt_get_slide_layouts_e2e/test_layouts.py --test 1
uv run python manual_tests/ppt_get_slide_layouts_e2e/test_layouts.py --test all
uv run python manual_tests/ppt_get_slide_layouts_e2e/test_layouts.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
