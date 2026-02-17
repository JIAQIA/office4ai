# PPT Slide Management E2E Tests

`ppt:add:slide`、`ppt:delete:slide`、`ppt:move:slide`、`ppt:goto:slide` 事件的端到端测试。

## 测试文件

### test_add_slide.py — 添加幻灯片

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 空白新增 | empty.pptx | 默认添加一张幻灯片 |
| 2 | 指定位置 | simple.pptx | 在 insertIndex=0 插入幻灯片 |
| 3 | 指定布局 | simple.pptx | 使用 'Blank' 布局添加幻灯片 |

### test_delete_slide.py — 删除幻灯片

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 删除中间幻灯片 | multi_slide.pptx | 删除 5 页 PPT 的第 2 张 (index=1) |
| 2 | 删除末尾幻灯片 | multi_slide.pptx | 删除最后一张幻灯片 |

### test_move_slide.py — 移动幻灯片

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 前移 | multi_slide.pptx | 将第 3 张移到第 1 的位置 |
| 2 | 后移 | multi_slide.pptx | 将第 1 张移到第 4 的位置 |

### test_goto_slide.py — 跳转幻灯片

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 跳到首页 | multi_slide.pptx | 跳转到 slideIndex=0 |
| 2 | 跳到末页 | multi_slide.pptx | 跳转到最后一张幻灯片 |

## 运行

```bash
# 添加幻灯片
uv run python manual_tests/ppt_slide_management_e2e/test_add_slide.py --test all

# 删除幻灯片
uv run python manual_tests/ppt_slide_management_e2e/test_delete_slide.py --test all

# 移动幻灯片
uv run python manual_tests/ppt_slide_management_e2e/test_move_slide.py --test all

# 跳转幻灯片
uv run python manual_tests/ppt_slide_management_e2e/test_goto_slide.py --test all
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
