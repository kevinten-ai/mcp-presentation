# Presentation Gen MCP Server

<p align="center">
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.10+-blue.svg" alt="Python 3.10+"></a>
  <a href="https://modelcontextprotocol.io/"><img src="https://img.shields.io/badge/MCP-compatible-green.svg" alt="MCP"></a>
  <img src="https://img.shields.io/badge/version-1.0.0-blue.svg" alt="Version 1.0.0">
</p>

<p align="center">
  AI-assisted presentation generation via MCP.<br>
  Create PowerPoint slides from structured content with templates, layouts, and styling.<br>
  Works with Claude Code, Claude Desktop, Cursor, and any MCP-compatible client.
</p>

<p align="center">
  <a href="README_CN.md">中文文档</a>
</p>

## Features

- **4 tools** — `create_slides`, `add_slide`, `export_to_pdf`, `create_thumbnail`
- **3 templates** — minimal (clean), corporate (professional), creative (bold)
- **4 layouts** — title_and_content, title_only, image_full, two_column
- **Real .pptx output** — fully editable in PowerPoint, Google Slides, Keynote
- **PDF export** — one-click conversion via LibreOffice
- **Thumbnail preview** — render first slide as PNG
- **Built-in guides** — MCP Resources with template docs and usage tips
- Auto-save with timestamps to output directory

## Architecture

```
User Prompt → AI Assistant (Claude / Cursor) → MCP Server → python-pptx
                                                   ↓
                                             Save .pptx / .pdf / .png
```

### How It Works

The server uses `python-pptx` to programmatically build PowerPoint files from structured JSON input. Each slide is constructed from shapes (textboxes, rectangles, images) on blank layouts, giving full control over positioning and styling.

| Tool | Input | Output | Engine |
|---|---|---|---|
| `create_slides` | Title + slides JSON + template | `.pptx` file | python-pptx |
| `add_slide` | Existing .pptx + slide data | Updated `.pptx` | python-pptx |
| `export_to_pdf` | `.pptx` file | `.pdf` file | LibreOffice CLI |
| `create_thumbnail` | `.pptx` file | `.png` image | Pillow |

## Quick Start

**1. Configure MCP**

<details>
<summary><b>Claude Code (CLI)</b></summary>

```bash
claude mcp add --transport stdio mcp-presentation \
  -- uv --directory /path/to/mcp-presentation run presentation-gen
```
</details>

<details>
<summary><b>Claude Desktop / Cursor (JSON config)</b></summary>

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

**2. Use it** — just ask your AI assistant:

```
"Create a 5-slide presentation about AI trends in 2026 using the corporate template"
```

The presentation will be saved to the `output/` directory as a .pptx file.

## Tools

### create_slides — Build a Full Presentation

Create a complete PowerPoint from structured content.

```
create_slides(
  title="Q1 2026 Report",
  slides=[
    {"title": "Overview", "content": "Revenue up 25%\nNew markets entered\nTeam grew to 50", "layout": "title_and_content"},
    {"title": "Key Metrics", "content": "Revenue: $10M|Customers: 500", "layout": "two_column"},
    {"title": "Thank You", "layout": "title_only"}
  ],
  template="corporate"
)
```

### add_slide — Append to Existing Presentation

Add a single slide to a .pptx file that already exists.

```
add_slide(
  pptx_path="/path/to/presentation.pptx",
  title="New Section",
  content="Additional content here",
  layout="title_and_content"
)
```

### export_to_pdf — Convert to PDF

Convert a .pptx file to PDF using LibreOffice.

```
export_to_pdf(pptx_path="/path/to/presentation.pptx")
```

> Requires LibreOffice installed. Install: `brew install --cask libreoffice` (macOS) or `sudo apt install libreoffice` (Ubuntu).

### create_thumbnail — Generate Preview Image

Render the first slide as a PNG thumbnail.

```
create_thumbnail(pptx_path="/path/to/presentation.pptx", width=1280)
```

## Templates

| Template | Style | Background | Title Style | Best for |
|---|---|---|---|---|
| `minimal` | Clean, modern | White | Dark text, light weight | Technical docs, reports |
| `corporate` | Professional | White + blue header | White on blue bar | Business decks, proposals |
| `creative` | Bold, colorful | Purple gradient | White, large bold | Pitches, creative briefs |

## Slide Layouts

| Layout | Description | Content format |
|---|---|---|
| `title_and_content` | Title at top, bullet content below | Regular text, `\n` for bullets |
| `title_only` | Large centered title + optional subtitle | Title only, content as subtitle |
| `image_full` | Full-bleed image with caption bar | Provide `image_path`, title as caption |
| `two_column` | Title + two columns side by side | Separate columns with `\|` |

## Example JSON Input

```json
{
  "title": "Project Alpha",
  "slides": [
    {
      "title": "Introduction",
      "content": "What is Project Alpha?\nWhy it matters\nTimeline overview",
      "layout": "title_and_content"
    },
    {
      "title": "Comparison",
      "content": "Before:\nSlow process\nManual steps\nHigh error rate|After:\nAutomated pipeline\n10x faster\n99.9% accuracy",
      "layout": "two_column"
    },
    {
      "title": "",
      "image_path": "/path/to/chart.png",
      "layout": "image_full"
    },
    {
      "title": "Questions?",
      "content": "team@example.com",
      "layout": "title_only"
    }
  ],
  "template": "corporate"
}
```

## MCP Resources

The server exposes built-in documentation that AI assistants can automatically read:

| Resource URI | Description |
|---|---|
| `guide://templates` | Template comparison, slide layouts, and example JSON |
| `guide://tools` | Tool usage, workflow, and tips |

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `PRESENTATION_OUTPUT_DIR` | No | `./output` | Directory to save generated files |

## Custom Output Directory

```bash
--env PRESENTATION_OUTPUT_DIR=/absolute/path/to/your/presentations
```

Files are saved with timestamps: `presentation_20260331_143022.pptx`.

## Troubleshooting

### Error Reference

| Error | Root Cause | Solution |
|---|---|---|
| `Missing required parameter: title` | No title provided to create_slides | Pass a `title` string |
| `Missing required parameter: slides` | No slides array provided | Pass a `slides` array of slide objects |
| `Presentation not found` | Invalid pptx_path | Check the file path exists |
| `Invalid slides JSON` | Malformed slides array | Ensure slides is a valid JSON array |

### PDF Export Errors

| Error | Root Cause | Solution |
|---|---|---|
| `LibreOffice is not installed` | soffice not in PATH | Install: `brew install --cask libreoffice` (macOS) or `sudo apt install libreoffice` (Ubuntu) |
| `LibreOffice conversion failed` | Corrupted .pptx or LibreOffice issue | Verify the .pptx opens correctly in PowerPoint |

### Thumbnail Errors

| Error | Root Cause | Solution |
|---|---|---|
| `Presentation has no slides` | Empty .pptx file | Ensure the presentation has at least one slide |
| Font rendering issues | System fonts not found | Install Helvetica (macOS) or DejaVu Sans (Linux) |

## Prerequisites

- **Python 3.10+**
- **[uv](https://docs.astral.sh/uv/)** — install with `curl -LsSf https://astral.sh/uv/install.sh | sh`
- **LibreOffice** (optional, for PDF export) — `brew install --cask libreoffice`

## Local Development

```bash
git clone https://github.com/kevinten-ai/mcp-presentation.git
cd mcp-presentation

# Install dependencies
uv sync

# Run the server directly
uv run presentation-gen
```

### Debug with MCP Inspector

```bash
npx @modelcontextprotocol/inspector uv --directory /path/to/mcp-presentation run presentation-gen
```

## Related Projects

- [mcp-image-gen](https://github.com/kevinten-ai/mcp-image-gen) — AI image generation MCP server (Gemini + Imagen)
- [mcp-video-gen](https://github.com/kevinten-ai/mcp-video-gen) — Multi-provider AI video generation MCP server
- [mcp-3d-gen](https://github.com/kevinten-ai/mcp-3d-gen) — AI 3D model generation MCP server

## License

MIT — see [LICENSE](LICENSE) for details.
