# UAT 场景：Word 工具注册验收 (word-tools)

## 测试目标

验证 21 个 Word MCP 工具在 MCP Inspector 中正确注册，参数 schema 符合预期。

## 验证清单

### 读取工具（6 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-01 | `word_get_selected_content` | 选中内容 | `document_uri` | - |
| W-02 | `word_get_visible_content` | 可见内容 | `document_uri` | - |
| W-03 | `word_get_selection` | 选区信息 | `document_uri` | - |
| W-04 | `word_get_document_structure` | 文档结构/大纲 | `document_uri` | - |
| W-05 | `word_get_document_stats` | 文档统计 | `document_uri` | - |
| W-06 | `word_get_styles` | 样式列表 | `document_uri` | - |

### 文本操作工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-07 | `word_insert_text` | 插入文本 | `document_uri`, `text` | `position`, `format` |
| W-08 | `word_append_text` | 追加文本 | `document_uri`, `text` | `style` |
| W-09 | `word_replace_text` | 替换文本 | `document_uri`, `search_text`, `replace_text` | `options` |
| W-10 | `word_replace_selection` | 替换选区 | `document_uri`, `text` | `format` |
| W-11 | `word_select_text` | 选中文本 | `document_uri`, `search_text` | - |

### 多媒体工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-12 | `word_insert_image` | 插入图片 | `document_uri`, `image_source` | `width`, `height` |
| W-13 | `word_insert_table` | 插入表格 | `document_uri`, `rows`, `columns` | `data`, `style` |
| W-14 | `word_insert_equation` | 插入公式 | `document_uri`, `equation` | - |
| W-15 | `word_insert_toc` | 插入目录 | `document_uri` | - |

### 导出工具（1 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-16 | `word_export_content` | 导出内容 | `document_uri` | `format`, `range` |

### 评论工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-17 | `word_get_comments` | 获取评论 | `document_uri` | - |
| W-18 | `word_insert_comment` | 插入评论 | `document_uri`, `text` | - |
| W-19 | `word_delete_comment` | 删除评论 | `document_uri`, `comment_id` | - |
| W-20 | `word_reply_comment` | 回复评论 | `document_uri`, `comment_id`, `text` | - |
| W-21 | `word_resolve_comment` | 解决评论 | `document_uri`, `comment_id` | - |

## 逐项验证要点

每个工具检查：

1. **存在性**：工具出现在 MCP Inspector Tools 列表中
2. **名称**：`word_` 前缀，snake_case 命名
3. **描述**：非空，能让用户理解工具用途
4. **`document_uri`**：必填，类型 string
5. **其他参数**：必填/可选与上表一致，类型合理
