# CLAUDE.md — mcp-presentation

## Project Overview

MCP server for AI-assisted presentation generation. Four tools: `create_slides`, `add_slide`, `export_to_pdf`, `create_thumbnail`. Produces real .pptx files from structured JSON input with built-in template styling.

## Architecture

- **Entry**: `src/presentation_gen/__init__.py` → `server.main()` via asyncio
- **Core**: `src/presentation_gen/server.py` — MCP tools, slide building logic, PDF export, thumbnail rendering
- **Templates**: `src/presentation_gen/templates.py` — three built-in templates (minimal, corporate, creative) with font/color/layout definitions

## Key Design Decisions

- Templates define all visual styling (fonts, colors, backgrounds) — the same slide data renders differently per template
- Slides are built on blank layouts using `python-pptx` shapes (textboxes, rectangles, images) for full control
- `two_column` layout uses `|` as column separator in the content string
- PDF export delegates to LibreOffice CLI (`soffice --headless`), returns clear install instructions if not available
- Thumbnail rendering uses Pillow to draw a simplified version of the first slide

## Development

```bash
uv sync                    # install deps
uv run presentation-gen    # run server
```

## Common Pitfalls

- `export_to_pdf` requires LibreOffice installed — the tool returns actionable install instructions if missing
- `image_full` layout requires a valid local file path for the image — remote URLs are not supported
- The `slides` parameter in `create_slides` can be a JSON array or a JSON string (auto-parsed)
- `create_thumbnail` is an approximation — complex shapes/gradients may not render perfectly
