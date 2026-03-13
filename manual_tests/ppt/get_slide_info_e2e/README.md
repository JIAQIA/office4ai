# PPT Get Slide Info E2E Tests

`ppt:get:slideInfo` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 单幻灯片信息 | empty.pptx | 获取空白 PPT 的基本信息，验证 slideCount >= 1 |
| 2 | 多幻灯片信息 | simple.pptx | 获取 3 页 PPT 的信息，双重验证幻灯片数量 |
| 3 | 幻灯片尺寸 | simple.pptx | 获取 PPT 信息，验证幻灯片宽高有效 |
| 4 | 指定幻灯片详情 | simple.pptx | 通过 slideIndex=1 获取第二张幻灯片的详情 |

## 运行

```bash
# 单个测试
uv run python manual_tests/ppt/get_slide_info_e2e/test_basic_info.py --test 1

# 全部测试
uv run python manual_tests/ppt/get_slide_info_e2e/test_basic_info.py --test all

# 列出测试
uv run python manual_tests/ppt/get_slide_info_e2e/test_basic_info.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
