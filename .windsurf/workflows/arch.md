---
description: 项目架构
---

当前项目是基于 MCP 协议实现一个管理并编辑 office 文档的 MCP Server

目前已经有一个相对完成的 ide4ai 的 MCP Server 实现，其具体代码在： @examples/ide4ai

当前项目的架构就是需要参考 ide4ai 实现完整的 office4ai 能力，原因如下：

1. 从功能上来讲，两个项目提供的工具是同质的，均是对文件的增删改查与Workspace的管理
2. 从细节上来讲
  a. ide4ai 提供的编辑工具更多地是针对「纯文本」，而 office4ai 提供的编辑工具面对的环境更多是多模态
  b. ide4ai 提供要求是准确性高（代码的逻辑结构严谨，缩进，闭包等均需要严格校验），而 office4ai 更多是内容丰富性强，可能随时需要绘图，扩写等
  c. ide4ai 需要提供比如 Terminal 这种辅助工具，方便执行工程性的校验，帮助完成代码编辑与功能开发。而 office4ai 可能更多需要配合比如 playwright 等工具搜索互联网完成更广的知识与信息的检索收集，完成文档的编写。

因此从架构与功能实现上来讲，二者是非常相近的。整体项目分为几个部分：

1. 基于 MCP 协议，实现 MCP Server
2. 实现在 office 的编辑与管理工具，提供细节增删改查的工具组合
3. 相较于 ide4ai, office4ai 不需要提供额外的大工具封装（比如 ide4ai 需要额外封装Terminal模块），因为 office4ai 所需的工具（浏览器）非常独立，而且有独立的MCP封装。

参考当前 ide4ai 的能力，来设计与实现 office4ai。当前项目使用 uv 管理虚拟环境，运行命令时使用 uv run