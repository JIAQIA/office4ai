# UAT 场景：Resource 注册验收 (resources)

## 测试目标

验证 3 个 MCP Resource 在 MCP Inspector 中正确注册，URI、元数据符合预期。

## 验证清单

| # | 资源 URI | 预期 name | 预期 mimeType | 描述关键词 |
|---|----------|-----------|---------------|-----------|
| R-01 | `window://office4ai` | Office 工作区 | `text/plain` | 根索引，文档连接数量 |
| R-02 | `window://office4ai/word` | Word 工作区 | `text/plain` | Word 文档聚合 |
| R-03 | `window://office4ai/ppt` | PPT 工作区 | `text/plain` | PPT 文档聚合 |

## 逐项验证要点

每个资源检查：

1. **存在性**：资源出现在 MCP Inspector Resources 列表中
2. **URI 格式**：符合 `window://office4ai[/category]` 模式
3. **name**：非空中文名称
4. **description**：非空，描述资源用途
5. **mimeType**：`text/plain`
