# Presentation Gen MCP Server

<p align="center">
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.10+-blue.svg" alt="Python 3.10+"></a>
  <a href="https://modelcontextprotocol.io/"><img src="https://img.shields.io/badge/MCP-compatible-green.svg" alt="MCP"></a>
  <img src="https://img.shields.io/badge/version-1.0.0-blue.svg" alt="Version 1.0.0">
</p>

<p align="center">
  AI 辅助演示文稿生成 MCP 服务器。<br>
  从结构化内容创建 PowerPoint 幻灯片，支持模板、布局和样式。<br>
  支持 Claude Code、Claude Desktop、Cursor 及所有 MCP 兼容客户端。
</p>

<p align="center">
  <a href="README.md">English</a>
</p>

## 特性

- **4 个工具** — `create_slides`（创建演示文稿）、`add_slide`（添加幻灯片）、`export_to_pdf`（导出 PDF）、`create_thumbnail`（生成缩略图）
- **3 套模板** — minimal（简约）、corporate（商务）、creative（创意）
- **4 种布局** — title_and_content、title_only、image_full、two_column
- **真正的 .pptx 输出** — 可在 PowerPoint、Google Slides、Keynote 中编辑
- **PDF 导出** — 通过 LibreOffice 一键转换
- **缩略图预览** — 将首页幻灯片渲染为 PNG 图片
- **内置指南** — MCP Resources 提供模板文档和使用技巧
- 使用微秒级时间戳自动命名，避免文件名碰撞
- 安全的自定义输出：扩展名必须匹配，且绝不覆盖已有文件
- 资源上限：最多 100 张内容页，缩略图宽度范围为 64–4096 像素

## 架构

```
用户提示词 → AI 助手（Claude / Cursor）→ MCP Server → python-pptx
                                            ↓
                                      保存 .pptx / .pdf / .png
```

### 工作原理

服务器使用 `python-pptx` 从结构化 JSON 输入程序化地构建 PowerPoint 文件。每张幻灯片由形状（文本框、矩形、图片）在空白布局上构建，实现对定位和样式的完全控制。

| 工具 | 输入 | 输出 | 引擎 |
|---|---|---|---|
| `create_slides` | 标题 + 幻灯片 JSON + 模板 | `.pptx` 文件 | python-pptx |
| `add_slide` | 已有 .pptx + 幻灯片数据 | 更新的 `.pptx` | python-pptx |
| `export_to_pdf` | `.pptx` 文件 | `.pdf` 文件 | LibreOffice CLI |
| `create_thumbnail` | `.pptx` 文件 | `.png` 图片 | Pillow |

## 快速开始

**1. 配置 MCP**

<details>
<summary><b>Claude Code（命令行）</b></summary>

```bash
claude mcp add --transport stdio mcp-presentation \
  -- uv --directory /path/to/mcp-presentation run presentation-gen
```
</details>

<details>
<summary><b>Claude Desktop / Cursor（JSON 配置）</b></summary>

```json
{
  "mcpServers": {
    "mcp-presentation": {
      "command": "uv",
      "args": ["--directory", "/path/to/mcp-presentation", "run", "presentation-gen"]
    }
  }
}
```
</details>

**2. 使用** — 直接用自然语言告诉 AI 助手：

```
"创建一个关于 2026 年 AI 趋势的 5 页演示文稿，使用商务模板"
```

演示文稿将保存到 `output/` 目录下的 .pptx 文件。

## 工具

### create_slides — 创建完整演示文稿

从结构化内容创建完整的 PowerPoint 文件。

```
create_slides(
  title="2026 年 Q1 报告",
  slides=[
    {"title": "概览", "content": "收入增长 25%\n进入新市场\n团队扩大到 50 人", "layout": "title_and_content"},
    {"title": "关键指标", "content": "收入：1000 万|客户数：500", "layout": "two_column"},
    {"title": "谢谢", "layout": "title_only"}
  ],
  template="corporate"
)
```

### add_slide — 追加幻灯片

向已有的 .pptx 文件追加一张幻灯片。

```
add_slide(
  pptx_path="/path/to/presentation.pptx",
  title="新章节",
  content="补充内容",
  layout="title_and_content"
)
```

### export_to_pdf — 导出为 PDF

通过 LibreOffice 将 .pptx 文件转换为 PDF。

```
export_to_pdf(pptx_path="/path/to/presentation.pptx")
```

> 需要安装 LibreOffice。安装命令：`brew install --cask libreoffice`（macOS）或 `sudo apt install libreoffice`（Ubuntu）。

### create_thumbnail — 生成预览图

将首页幻灯片渲染为 PNG 缩略图。

```
create_thumbnail(pptx_path="/path/to/presentation.pptx", width=1280)
```

## 模板

| 模板 | 风格 | 背景 | 标题样式 | 适用场景 |
|---|---|---|---|---|
| `minimal` | 简约现代 | 白色 | 深色文字，细字体 | 技术文档、报告 |
| `corporate` | 专业商务 | 白色 + 蓝色头部栏 | 蓝底白字 | 商业演示、方案 |
| `creative` | 大胆多彩 | 紫色渐变 | 白色粗体大字 | 路演、创意简报 |

## 幻灯片布局

| 布局 | 描述 | 内容格式 |
|---|---|---|
| `title_and_content` | 顶部标题，下方要点内容 | 普通文本，`\n` 分隔要点 |
| `title_only` | 大号居中标题 + 可选副标题 | 仅标题，content 作为副标题 |
| `image_full` | 全幅图片 + 底部说明栏 | 提供 `image_path`，title 作为说明 |
| `two_column` | 标题 + 左右两列内容 | 用 `\|` 分隔左右列 |

## 示例 JSON 输入

```json
{
  "title": "Alpha 项目",
  "slides": [
    {
      "title": "简介",
      "content": "什么是 Alpha 项目？\n为什么重要\n时间线概览",
      "layout": "title_and_content"
    },
    {
      "title": "对比",
      "content": "改造前：\n流程缓慢\n手动操作\n错误率高|改造后：\n自动化流水线\n速度提升 10 倍\n准确率 99.9%",
      "layout": "two_column"
    },
    {
      "title": "",
      "image_path": "/path/to/chart.png",
      "layout": "image_full"
    },
    {
      "title": "提问环节",
      "content": "team@example.com",
      "layout": "title_only"
    }
  ],
  "template": "corporate"
}
```

## MCP Resources

服务器内置文档，AI 助手可自动读取参考：

| 资源 URI | 说明 |
|---|---|
| `guide://templates` | 模板对比、幻灯片布局、示例 JSON |
| `guide://tools` | 工具用法、工作流程、使用技巧 |

## 环境变量

| 变量 | 必填 | 默认值 | 说明 |
|---|---|---|---|
| `PRESENTATION_OUTPUT_DIR` | 否 | `./output` | 文件保存目录 |

## 自定义输出目录

```bash
--env PRESENTATION_OUTPUT_DIR=/你的/演示文稿/保存/路径
```

文件以微秒级时间戳命名保存：`presentation_20260331_143022_123456.pptx`。
自定义输出路径必须使用对应格式的扩展名（`.pptx`、`.pdf` 或 `.png`），
并且目标文件不能已经存在。`add_slide` 是唯一的原位修改操作，会按设计更新传入的 `.pptx` 文件。

## 常见问题排查

### 错误速查

| 错误 | 根因 | 解决方案 |
|---|---|---|
| `Missing required parameter: title` | 未提供标题 | 传入 `title` 字符串 |
| `Missing required parameter: slides` | 未提供幻灯片数组 | 传入 `slides` 幻灯片对象数组 |
| `Presentation not found` | pptx_path 路径无效 | 检查文件路径是否存在 |
| `Invalid slides JSON` | 幻灯片数组格式错误 | 确保 slides 是有效的 JSON 数组 |
| `Output file already exists` | 输出会覆盖已有文件 | 选择新路径，或由用户明确删除旧文件 |
| `output_path must use...` | 输出扩展名与生成格式不符 | 按工具使用 `.pptx`、`.pdf` 或 `.png` |
| `slides must contain at most 100 items` | 单次演示文稿过大 | 拆分为多个演示文稿 |

### PDF 导出错误

| 错误 | 根因 | 解决方案 |
|---|---|---|
| `LibreOffice is not installed` | 系统未安装 LibreOffice | 安装：`brew install --cask libreoffice`（macOS）或 `sudo apt install libreoffice`（Ubuntu）|
| `LibreOffice conversion failed` | .pptx 文件损坏或 LibreOffice 异常 | 确认 .pptx 文件可在 PowerPoint 中正常打开 |

### 缩略图错误

| 错误 | 根因 | 解决方案 |
|---|---|---|
| `Presentation has no slides` | .pptx 文件为空 | 确保演示文稿至少有一张幻灯片 |
| 字体渲染异常 | 系统缺少字体 | 安装 Helvetica（macOS）或 DejaVu Sans（Linux）|

## 前置要求

- **Python 3.10+**
- **[uv](https://docs.astral.sh/uv/)** — 安装命令：`curl -LsSf https://astral.sh/uv/install.sh | sh`
- **LibreOffice**（可选，用于 PDF 导出）— `brew install --cask libreoffice`

## 本地开发

```bash
git clone https://github.com/kevinten-ai/mcp-presentation.git
cd mcp-presentation

# 安装依赖
uv sync

# 直接运行服务器
uv run presentation-gen
```

### 使用 MCP Inspector 调试

```bash
npx @modelcontextprotocol/inspector uv --directory /path/to/mcp-presentation run presentation-gen
```

## 相关媒体工作流

图片生成、图片编辑和视频生成默认使用 AnyCap CLI。`mcp-image-gen` 与
`mcp-video-gen` 仅保留用于 MCP 兼容性和协议测试，不作为日常生成入口。
本服务器继续负责代理工作流中的本地 PPTX、PDF 和缩略图操作。

## 许可证

MIT — 详见 [LICENSE](LICENSE)。
