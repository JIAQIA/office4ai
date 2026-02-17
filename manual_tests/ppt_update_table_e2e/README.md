# PPT Update Table E2E Tests

`ppt:update:tableCell`、`ppt:update:tableRowColumn`、`ppt:update:tableFormat` 事件的端到端测试。

## 测试文件

### test_cell_update.py — 单元格更新

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 单格更新 | empty.pptx | 更新表格 [0,0] 单元格 |
| 2 | 多格更新 | empty.pptx | 批量更新对角线单元格 |
| 3 | 边角格更新 | empty.pptx | 更新表格四角单元格 |

### test_row_column_update.py — 行列更新

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 更新行 | empty.pptx | 更新表格第一行为 H1/H2/H3 |
| 2 | 更新列 | empty.pptx | 更新表格第一列为 R1/R2/R3 |
| 3 | 混合更新 | empty.pptx | 同时更新第一行和最后一列 |

### test_table_format.py — 表格格式

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 单元格格式 | empty.pptx | 设置 [0,0] 背景红色+粗体 |
| 2 | 行格式 | empty.pptx | 设置第一行绿色背景 |
| 3 | 列格式 | empty.pptx | 设置第一列蓝色背景 |

## 运行

```bash
# 单元格更新
uv run python manual_tests/ppt_update_table_e2e/test_cell_update.py --test all

# 行列更新
uv run python manual_tests/ppt_update_table_e2e/test_row_column_update.py --test all

# 表格格式
uv run python manual_tests/ppt_update_table_e2e/test_table_format.py --test all
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
