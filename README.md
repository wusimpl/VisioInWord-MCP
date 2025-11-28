# Word-Visio MCP Server

一个 MCP (Model Context Protocol) 服务器，让大模型可以直接在 Word 文档中插入 Visio UML 类图和文本。

## 功能

- insert_text: 在 Word 文档光标位置插入纯文本
- insert_uml_class: 在 Word 文档光标位置插入 UML 类图（通过 Visio 绘制）

## 系统要求

- Windows 操作系统
- Microsoft Word（已安装并可运行）
- Microsoft Visio（已安装并可运行）
- Python 3.10+

## 安装

1. 下载本项目

2. 安装依赖：

```bash
pip install mcp pywin32
```

## 在 Claude Code 中配置

使用 stdio 传输方式添加本地 MCP 服务器：

```bash
claude mcp add --transport stdio word-visio -- python /path/to/word_visio_mcp.py
```

请将 `/path/to/word_visio_mcp.py` 替换为实际的脚本绝对路径。

## 使用方法

1. 打开 Microsoft Word 并创建或打开一个文档
2. 将光标放置在要插入内容的位置
3. 在 Claude Code 中告诉模型调用word-visio mcp生成类图

## 注意事项

- 使用前必须先打开 Word 并将光标放置在目标位置
- 插入类图时会短暂打开 Visio 窗口，完成后自动关闭
- 如果遇到 COM 错误，请确保 Word 和 Visio 已正确安装并激活

## 故障排除

1. "请确保 Word 已打开且有活动文档" - 请先打开 Word 并创建/打开文档
2. Visio 相关错误 - 确认 Visio 已安装，尝试手动打开 Visio 确认可正常运行
3. pywin32 错误 - 重新安装 pywin32: `pip install --upgrade pywin32`

## 许可证

MIT License
