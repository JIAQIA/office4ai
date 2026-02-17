# PPT Insert Text E2E Tests

`ppt:insert:text` 事件的端到端测试。

## 测试文件

### test_basic_insert.py — 基本文本插入

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 简单文本插入 | empty.pptx | 在空白幻灯片插入 'Hello PPT' |
| 2 | 带位置插入 | empty.pptx | 在指定位置 (100, 100) 插入文本 |
| 3 | 带字体插入 | empty.pptx | 插入带 fontSize=24, fontName=Arial 的文本 |
| 4 | 指定幻灯片插入 | simple.pptx | 在 slideIndex=1 (第二张) 上插入文本 |

### test_insert_options.py — 高级插入选项

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 带填充色 | empty.pptx | 插入带黄色背景的文本框 |
| 2 | 连续多个文本框 | empty.pptx | 在同一幻灯片连续插入两个文本框 |
| 3 | 中文文本插入 | empty.pptx | 插入较长的中文文本 |

## 运行

```bash
# 基本插入
uv run python manual_tests/ppt_insert_text_e2e/test_basic_insert.py --test all

# 高级选项
uv run python manual_tests/ppt_insert_text_e2e/test_insert_options.py --test all
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
