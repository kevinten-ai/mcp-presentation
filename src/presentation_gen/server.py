import asyncio
import json
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

from mcp.server.models import InitializationOptions
import mcp.types as types
from mcp.server import NotificationOptions, Server
import mcp.server.stdio

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from PIL import Image, ImageDraw, ImageFont

from .templates import (
    TemplateStyle,
    get_template,
    TEMPLATES,
    LAYOUTS,
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
)

# ── Configuration ────────────────────────────────────────────────────────────
OUTPUT_DIR = os.getenv("PRESENTATION_OUTPUT_DIR", os.path.join(os.getcwd(), "output"))

# ── Resource content ─────────────────────────────────────────────────────────
TEMPLATE_GUIDE = """# Presentation Template Guide

## Available Templates

| Template | Style | Background | Title | Best for |
|---|---|---|---|---|
| `minimal` | Clean, modern | White | Dark text, light font | Technical docs, reports |
| `corporate` | Professional | White + blue header bar | White on blue | Business decks, proposals |
| `creative` | Bold, colorful | Purple gradient | White, large | Pitches, creative briefs |

## Slide Layouts

| Layout | Description | Use case |
|---|---|---|
| `title_and_content` | Title at top, bullet content below | Standard information slides |
| `title_only` | Large centered title | Section dividers, opening slides |
| `image_full` | Full-bleed image with optional caption | Photo showcases, visual slides |
| `two_column` | Title + two content columns | Comparisons, pros/cons |

## Example JSON Input

```json
{
  "title": "Q1 Report",
  "slides": [
    {"title": "Overview", "content": "Key metrics and highlights", "layout": "title_and_content"},
    {"title": "Revenue Growth", "content": "Left column|Right column", "layout": "two_column"},
    {"title": "", "image_path": "/path/to/chart.png", "layout": "image_full"}
  ],
  "template": "corporate"
}
```

## Tips

- For `two_column` layout, separate left and right content with `|`
- For `image_full`, provide `image_path` — the image will fill the slide
- The `content` field supports line breaks (`\\n`) for bullet points
- Templates define fonts, colors, and styling — the same slide data looks different with each template
"""

TOOLS_GUIDE = """# Presentation Tools Guide

## Tools Overview

| Tool | Description | Key params |
|---|---|---|
| `create_slides` | Create a full presentation from structured JSON | `title`, `slides[]`, `template` |
| `add_slide` | Append a single slide to an existing .pptx | `pptx_path`, `title`, `content` |
| `export_to_pdf` | Convert .pptx to PDF via LibreOffice | `pptx_path` |
| `create_thumbnail` | Render first slide as a PNG thumbnail | `pptx_path`, `width` |

## Workflow

1. Use `create_slides` to build the initial presentation
2. Use `add_slide` to append more slides incrementally
3. Use `export_to_pdf` to produce a PDF version
4. Use `create_thumbnail` to generate a preview image

## Output

All files are saved to the configured output directory (default: `./output/`).
Files are timestamped to avoid collisions: `presentation_20260331_143022.pptx`.
"""

# ── Server ───────────────────────────────────────────────────────────────────
server = Server("mcp-presentation")


# ── Slide building helpers ───────────────────────────────────────────────────

def _apply_background(slide, style: TemplateStyle) -> None:
    """Apply background color or gradient to a slide."""
    background = slide.background
    fill = background.fill

    if style.has_gradient_bg:
        fill.gradient()
        fill.gradient_stops[0].color.rgb = style.gradient_color_start
        fill.gradient_stops[0].position = 0.0
        fill.gradient_stops[1].color.rgb = style.gradient_color_end
        fill.gradient_stops[1].position = 1.0
    else:
        fill.solid()
        fill.fore_color.rgb = style.bg_color


def _add_header_bar(slide, style: TemplateStyle) -> None:
    """Add a colored header bar across the top of the slide (corporate template)."""
    if not style.has_header_bar:
        return
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Emu(0),
        top=Emu(0),
        width=SLIDE_WIDTH,
        height=style.header_bar_height,
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = style.accent_color
    shape.line.fill.background()


def _add_text_box(
    slide,
    text: str,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    font_name: str = "Calibri",
    font_size: Pt = Pt(18),
    font_color: RGBColor = RGBColor(0x33, 0x33, 0x33),
    bold: bool = False,
    alignment: PP_ALIGN = PP_ALIGN.LEFT,
    vertical_anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
) -> None:
    """Add a styled text box to a slide."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    # Vertical centering
    try:
        tf.paragraphs[0].alignment = alignment
    except Exception:
        pass

    # Set vertical anchor on the text frame
    txBox.text_frame._txBody.bodyPr.set("anchor", {
        MSO_ANCHOR.TOP: "t",
        MSO_ANCHOR.MIDDLE: "ctr",
        MSO_ANCHOR.BOTTOM: "b",
    }.get(vertical_anchor, "t"))

    # Parse content: support newlines as separate paragraphs (bullet points)
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line.strip()
        p.font.name = font_name
        p.font.size = font_size
        p.font.color.rgb = font_color
        p.font.bold = bold
        p.alignment = alignment


def _build_title_and_content(slide, slide_data: dict, style: TemplateStyle) -> None:
    """Layout: title at top, bullet content below."""
    _apply_background(slide, style)

    if style.has_header_bar:
        _add_header_bar(slide, style)
        # Title inside header bar
        _add_text_box(
            slide,
            slide_data.get("title", ""),
            left=Inches(0.8),
            top=Inches(0.15),
            width=Inches(11.7),
            height=Inches(0.9),
            font_name=style.title_font,
            font_size=style.title_size,
            font_color=style.title_color,
            bold=style.title_bold,
            alignment=PP_ALIGN.LEFT,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )
        content_top = Inches(1.5)
    else:
        # Title at top
        _add_text_box(
            slide,
            slide_data.get("title", ""),
            left=Inches(0.8),
            top=Inches(0.5),
            width=Inches(11.7),
            height=Inches(1.0),
            font_name=style.title_font,
            font_size=style.title_size,
            font_color=style.title_color,
            bold=style.title_bold,
            alignment=PP_ALIGN.LEFT,
        )
        content_top = Inches(1.8)

    # Body content
    content = slide_data.get("content", "")
    if content:
        _add_text_box(
            slide,
            content,
            left=Inches(0.8),
            top=content_top,
            width=Inches(11.7),
            height=Inches(5.0),
            font_name=style.body_font,
            font_size=style.body_size,
            font_color=style.body_color,
        )

    # Optional image below content
    image_path = slide_data.get("image_path")
    if image_path and Path(image_path).exists():
        try:
            slide.shapes.add_picture(
                image_path,
                left=Inches(3.5),
                top=Inches(4.0),
                width=Inches(6.0),
            )
        except Exception:
            pass


def _build_title_only(slide, slide_data: dict, style: TemplateStyle) -> None:
    """Layout: large centered title."""
    _apply_background(slide, style)

    if style.has_header_bar:
        # Full-height accent background for title-only
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left=Emu(0),
            top=Emu(0),
            width=SLIDE_WIDTH,
            height=SLIDE_HEIGHT,
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = style.accent_color
        shape.line.fill.background()

    _add_text_box(
        slide,
        slide_data.get("title", ""),
        left=Inches(1.5),
        top=Inches(2.5),
        width=Inches(10.3),
        height=Inches(2.5),
        font_name=style.title_font,
        font_size=Pt(48),
        font_color=style.title_color if style.has_header_bar or style.has_gradient_bg
        else style.title_color,
        bold=style.title_bold,
        alignment=PP_ALIGN.CENTER,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )

    # Optional subtitle from content
    content = slide_data.get("content", "")
    if content:
        subtitle_color = (
            RGBColor(0xCC, 0xCC, 0xCC)
            if (style.has_header_bar or style.has_gradient_bg)
            else style.body_color
        )
        _add_text_box(
            slide,
            content,
            left=Inches(2.0),
            top=Inches(4.5),
            width=Inches(9.3),
            height=Inches(1.5),
            font_name=style.body_font,
            font_size=Pt(24),
            font_color=subtitle_color,
            alignment=PP_ALIGN.CENTER,
        )


def _build_image_full(slide, slide_data: dict, style: TemplateStyle) -> None:
    """Layout: full-bleed image with optional caption."""
    _apply_background(slide, style)

    image_path = slide_data.get("image_path")
    if image_path and Path(image_path).exists():
        try:
            slide.shapes.add_picture(
                image_path,
                left=Emu(0),
                top=Emu(0),
                width=SLIDE_WIDTH,
                height=SLIDE_HEIGHT,
            )
        except Exception:
            # If image fails, show placeholder text
            _add_text_box(
                slide,
                f"[Image not loaded: {image_path}]",
                left=Inches(2.0),
                top=Inches(3.0),
                width=Inches(9.3),
                height=Inches(1.5),
                font_color=style.body_color,
                alignment=PP_ALIGN.CENTER,
            )
    else:
        _add_text_box(
            slide,
            f"[Image: {image_path or 'not specified'}]",
            left=Inches(2.0),
            top=Inches(3.0),
            width=Inches(9.3),
            height=Inches(1.5),
            font_color=style.body_color,
            alignment=PP_ALIGN.CENTER,
        )

    # Caption at bottom
    title = slide_data.get("title", "")
    if title:
        # Semi-transparent bar at bottom for caption
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left=Emu(0),
            top=Inches(6.2),
            width=SLIDE_WIDTH,
            height=Inches(1.3),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
        # Set transparency (50%)
        bar.fill.fore_color.brightness = 0.0
        bar.line.fill.background()

        _add_text_box(
            slide,
            title,
            left=Inches(0.8),
            top=Inches(6.35),
            width=Inches(11.7),
            height=Inches(1.0),
            font_name=style.title_font,
            font_size=Pt(24),
            font_color=RGBColor(0xFF, 0xFF, 0xFF),
            bold=True,
            alignment=PP_ALIGN.LEFT,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )


def _build_two_column(slide, slide_data: dict, style: TemplateStyle) -> None:
    """Layout: title plus two content columns side by side."""
    _apply_background(slide, style)

    if style.has_header_bar:
        _add_header_bar(slide, style)
        _add_text_box(
            slide,
            slide_data.get("title", ""),
            left=Inches(0.8),
            top=Inches(0.15),
            width=Inches(11.7),
            height=Inches(0.9),
            font_name=style.title_font,
            font_size=style.title_size,
            font_color=style.title_color,
            bold=style.title_bold,
            alignment=PP_ALIGN.LEFT,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )
        content_top = Inches(1.5)
    else:
        _add_text_box(
            slide,
            slide_data.get("title", ""),
            left=Inches(0.8),
            top=Inches(0.5),
            width=Inches(11.7),
            height=Inches(1.0),
            font_name=style.title_font,
            font_size=style.title_size,
            font_color=style.title_color,
            bold=style.title_bold,
            alignment=PP_ALIGN.LEFT,
        )
        content_top = Inches(1.8)

    # Split content by | for two columns
    content = slide_data.get("content", "")
    if "|" in content:
        parts = content.split("|", 1)
        left_text = parts[0].strip()
        right_text = parts[1].strip()
    else:
        left_text = content
        right_text = ""

    # Divider line
    divider = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(6.55),
        top=content_top,
        width=Inches(0.03),
        height=Inches(5.0),
    )
    divider.fill.solid()
    divider.fill.fore_color.rgb = style.accent_color
    divider.line.fill.background()

    # Left column
    _add_text_box(
        slide,
        left_text,
        left=Inches(0.8),
        top=content_top,
        width=Inches(5.5),
        height=Inches(5.0),
        font_name=style.body_font,
        font_size=style.body_size,
        font_color=style.body_color,
    )

    # Right column
    _add_text_box(
        slide,
        right_text,
        left=Inches(6.9),
        top=content_top,
        width=Inches(5.5),
        height=Inches(5.0),
        font_name=style.body_font,
        font_size=style.body_size,
        font_color=style.body_color,
    )


# Layout dispatcher
LAYOUT_BUILDERS = {
    "title_and_content": _build_title_and_content,
    "title_only": _build_title_only,
    "image_full": _build_image_full,
    "two_column": _build_two_column,
}


def _add_slide_to_prs(prs: Presentation, slide_data: dict, style: TemplateStyle) -> None:
    """Add a single slide to a Presentation object using the specified layout and style."""
    layout_name = slide_data.get("layout", "title_and_content")
    builder = LAYOUT_BUILDERS.get(layout_name, _build_title_and_content)

    # Use blank layout
    blank_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_layout)

    builder(slide, slide_data, style)


def _create_presentation(
    title: str,
    slides: list[dict],
    template: str = "minimal",
    output_path: str | None = None,
) -> str:
    """Create a complete PowerPoint presentation and save to disk."""
    style = get_template(template)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # Title slide (always first)
    title_slide_data = {"title": title, "content": "", "layout": "title_only"}
    _add_slide_to_prs(prs, title_slide_data, style)

    # Content slides
    for slide_data in slides:
        _add_slide_to_prs(prs, slide_data, style)

    # Determine output path
    if output_path:
        filepath = Path(output_path)
    else:
        output_dir = Path(OUTPUT_DIR)
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = output_dir / f"presentation_{timestamp}.pptx"

    filepath.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(filepath))
    return str(filepath)


def _add_slide_to_existing(
    pptx_path: str,
    title: str,
    content: str = "",
    image_path: str | None = None,
    layout: str = "title_and_content",
    template: str = "minimal",
) -> str:
    """Add a single slide to an existing presentation."""
    path = Path(pptx_path)
    if not path.exists():
        raise FileNotFoundError(f"Presentation not found: {pptx_path}")

    prs = Presentation(str(path))
    # Try to detect template from existing slides, default to provided
    style = get_template(template)

    slide_data = {
        "title": title,
        "content": content,
        "layout": layout,
    }
    if image_path:
        slide_data["image_path"] = image_path

    _add_slide_to_prs(prs, slide_data, style)
    prs.save(str(path))
    return str(path)


def _export_to_pdf(pptx_path: str, output_path: str | None = None) -> str:
    """Convert PPTX to PDF using LibreOffice CLI."""
    path = Path(pptx_path)
    if not path.exists():
        raise FileNotFoundError(f"Presentation not found: {pptx_path}")

    # Check for LibreOffice
    soffice = shutil.which("soffice")
    if not soffice:
        # Also check common macOS path
        macos_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if Path(macos_path).exists():
            soffice = macos_path

    if not soffice:
        raise RuntimeError(
            "LibreOffice is not installed or not in PATH. "
            "Install it to enable PDF export:\n"
            "  macOS:  brew install --cask libreoffice\n"
            "  Ubuntu: sudo apt install libreoffice\n"
            "  Windows: Download from https://www.libreoffice.org/download/"
        )

    if output_path:
        out_dir = str(Path(output_path).parent)
    else:
        out_dir = str(path.parent)

    # Run LibreOffice conversion
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, str(path)],
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    # LibreOffice outputs to same directory with .pdf extension
    generated_pdf = Path(out_dir) / f"{path.stem}.pdf"

    if output_path and str(generated_pdf) != output_path:
        generated_pdf.rename(output_path)
        return output_path

    return str(generated_pdf)


def _create_thumbnail(
    pptx_path: str,
    width: int = 1280,
    output_path: str | None = None,
) -> str:
    """Generate a thumbnail image from the first slide of a PPTX."""
    path = Path(pptx_path)
    if not path.exists():
        raise FileNotFoundError(f"Presentation not found: {pptx_path}")

    prs = Presentation(str(path))
    if len(prs.slides) == 0:
        raise ValueError("Presentation has no slides")

    slide = prs.slides[0]
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    # Calculate pixel dimensions (maintain aspect ratio)
    aspect = slide_h / slide_w
    height = int(width * aspect)

    # Create canvas
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Scale factor: EMU to pixels
    scale_x = width / slide_w
    scale_y = height / slide_h

    # Render background — check for fill
    bg = slide.background
    try:
        bg_fill = bg.fill
        if bg_fill.type is not None:
            # Attempt to get solid color
            try:
                rgb = bg_fill.fore_color.rgb
                bg_color = (rgb[0], rgb[1], rgb[2])
                draw.rectangle([(0, 0), (width, height)], fill=bg_color)
            except Exception:
                pass
    except Exception:
        pass

    # Render shapes
    for shape in slide.shapes:
        left = int(shape.left * scale_x)
        top = int(shape.top * scale_y)
        right = int((shape.left + shape.width) * scale_x)
        bottom = int((shape.top + shape.height) * scale_y)

        # Render filled rectangles (header bars, backgrounds)
        if shape.shape_type and hasattr(shape, "fill"):
            try:
                fill = shape.fill
                if fill.type is not None:
                    try:
                        rgb = fill.fore_color.rgb
                        color = (rgb[0], rgb[1], rgb[2])
                        draw.rectangle([(left, top), (right, bottom)], fill=color)
                    except Exception:
                        pass
            except Exception:
                pass

        # Render images
        if shape.shape_type == 13:  # Picture
            try:
                image_blob = shape.image.blob
                from io import BytesIO
                pic = Image.open(BytesIO(image_blob))
                pic = pic.resize((right - left, bottom - top), Image.LANCZOS)
                img.paste(pic, (left, top))
            except Exception:
                pass

        # Render text
        if shape.has_text_frame:
            tf = shape.text_frame
            y_offset = top
            for paragraph in tf.paragraphs:
                text = paragraph.text.strip()
                if not text:
                    y_offset += 20
                    continue

                # Determine font size in pixels
                font_size_pt = 18  # default
                if paragraph.font and paragraph.font.size:
                    font_size_pt = int(paragraph.font.size.pt)
                elif len(paragraph.runs) > 0 and paragraph.runs[0].font.size:
                    font_size_pt = int(paragraph.runs[0].font.size.pt)

                pixel_size = int(font_size_pt * (width / 960))

                # Try to load a decent font, fall back to default
                try:
                    font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", pixel_size)
                except Exception:
                    try:
                        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", pixel_size)
                    except Exception:
                        font = ImageFont.load_default()

                # Get text color
                text_color = (51, 51, 51)  # default dark gray
                try:
                    if paragraph.font and paragraph.font.color and paragraph.font.color.rgb:
                        rgb = paragraph.font.color.rgb
                        text_color = (rgb[0], rgb[1], rgb[2])
                    elif len(paragraph.runs) > 0:
                        run = paragraph.runs[0]
                        if run.font.color and run.font.color.rgb:
                            rgb = run.font.color.rgb
                            text_color = (rgb[0], rgb[1], rgb[2])
                except Exception:
                    pass

                # Get text bounding box for alignment
                bbox = draw.textbbox((0, 0), text, font=font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]

                # Alignment
                x = left + 5
                try:
                    align = paragraph.alignment
                    if align == PP_ALIGN.CENTER:
                        x = left + (right - left - text_w) // 2
                    elif align == PP_ALIGN.RIGHT:
                        x = right - text_w - 5
                except Exception:
                    pass

                draw.text((x, y_offset), text, fill=text_color, font=font)
                y_offset += text_h + 8

    # Save
    if output_path:
        filepath = Path(output_path)
    else:
        output_dir = Path(OUTPUT_DIR)
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = output_dir / f"thumbnail_{timestamp}.png"

    filepath.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(filepath), "PNG")
    return str(filepath)


# ── Resources ────────────────────────────────────────────────────────────────

@server.list_resources()
async def handle_list_resources() -> list[types.Resource]:
    return [
        types.Resource(
            uri="guide://templates",
            name="Template & Layout Guide",
            description="Available templates, slide layouts, and example JSON input",
            mimeType="text/markdown",
        ),
        types.Resource(
            uri="guide://tools",
            name="Tools Usage Guide",
            description="How to use each presentation tool, workflow, and tips",
            mimeType="text/markdown",
        ),
    ]


@server.read_resource()
async def handle_read_resource(uri: types.AnyUrl) -> str:
    uri_str = str(uri)
    if uri_str == "guide://templates":
        return TEMPLATE_GUIDE
    elif uri_str == "guide://tools":
        return TOOLS_GUIDE
    raise ValueError(f"Unknown resource: {uri_str}")


# ── Tools ────────────────────────────────────────────────────────────────────

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    template_names = list(TEMPLATES.keys())
    layout_names = list(LAYOUTS.keys())

    return [
        types.Tool(
            name="create_slides",
            description=(
                "Create a PowerPoint presentation from structured content. "
                f"Templates: {', '.join(template_names)}. "
                f"Layouts: {', '.join(layout_names)}. "
                "Pass a title and a JSON array of slide objects."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "Presentation title (shown on the first title slide)",
                    },
                    "slides": {
                        "type": "array",
                        "description": (
                            "Array of slide objects. Each slide: "
                            '{"title": "...", "content": "...", "image_path": "..." (optional), '
                            '"layout": "title_and_content"/"title_only"/"image_full"/"two_column"}'
                        ),
                        "items": {
                            "type": "object",
                            "properties": {
                                "title": {"type": "string", "description": "Slide title"},
                                "content": {"type": "string", "description": "Slide body text. Use \\n for bullet points. For two_column layout, use | to separate columns."},
                                "image_path": {"type": "string", "description": "Path to an image file (optional)"},
                                "layout": {
                                    "type": "string",
                                    "description": "Slide layout",
                                    "enum": layout_names,
                                    "default": "title_and_content",
                                },
                            },
                            "required": ["title"],
                        },
                    },
                    "template": {
                        "type": "string",
                        "description": "Visual template to use",
                        "enum": template_names,
                        "default": "minimal",
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Custom output file path (optional, defaults to output/ directory with timestamp)",
                    },
                },
                "required": ["title", "slides"],
            },
        ),
        types.Tool(
            name="add_slide",
            description=(
                "Add a single slide to an existing PowerPoint presentation. "
                f"Layouts: {', '.join(layout_names)}."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Path to the existing .pptx file",
                    },
                    "title": {
                        "type": "string",
                        "description": "Slide title",
                    },
                    "content": {
                        "type": "string",
                        "description": "Slide body text (optional). Use \\n for bullet points.",
                    },
                    "image_path": {
                        "type": "string",
                        "description": "Path to an image file (optional)",
                    },
                    "layout": {
                        "type": "string",
                        "description": "Slide layout",
                        "enum": layout_names,
                        "default": "title_and_content",
                    },
                },
                "required": ["pptx_path", "title"],
            },
        ),
        types.Tool(
            name="export_to_pdf",
            description=(
                "Convert a PowerPoint file to PDF using LibreOffice. "
                "Requires LibreOffice installed on the system."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Path to the .pptx file to convert",
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Custom output PDF path (optional, defaults to same directory as input)",
                    },
                },
                "required": ["pptx_path"],
            },
        ),
        types.Tool(
            name="create_thumbnail",
            description=(
                "Generate a thumbnail PNG image from the first slide of a presentation. "
                "Renders shapes, text, and images using Pillow."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Path to the .pptx file",
                    },
                    "width": {
                        "type": "integer",
                        "description": "Thumbnail width in pixels (default 1280, height auto-calculated from 16:9)",
                        "default": 1280,
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Custom output PNG path (optional, defaults to output/ directory with timestamp)",
                    },
                },
                "required": ["pptx_path"],
            },
        ),
    ]


@server.call_tool()
async def handle_call_tool(
    name: str, arguments: dict | None
) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
    if not arguments:
        return [types.TextContent(type="text", text="Missing arguments")]

    if name == "create_slides":
        title = arguments.get("title")
        slides = arguments.get("slides")
        if not title:
            return [types.TextContent(type="text", text="Missing required parameter: title")]
        if not slides:
            return [types.TextContent(type="text", text="Missing required parameter: slides")]

        # Parse slides if passed as string
        if isinstance(slides, str):
            try:
                slides = json.loads(slides)
            except json.JSONDecodeError as e:
                return [types.TextContent(type="text", text=f"Invalid slides JSON: {e}")]

        template = arguments.get("template", "minimal")
        output_path = arguments.get("output_path")

        try:
            filepath = _create_presentation(title, slides, template, output_path)
            slide_count = len(slides) + 1  # +1 for auto-generated title slide
            return [types.TextContent(
                type="text",
                text=f"Presentation created successfully!\n"
                     f"  File: {filepath}\n"
                     f"  Slides: {slide_count}\n"
                     f"  Template: {template}",
            )]
        except Exception as e:
            return [types.TextContent(type="text", text=f"Error creating presentation: {e}")]

    if name == "add_slide":
        pptx_path = arguments.get("pptx_path")
        title = arguments.get("title")
        if not pptx_path:
            return [types.TextContent(type="text", text="Missing required parameter: pptx_path")]
        if not title:
            return [types.TextContent(type="text", text="Missing required parameter: title")]

        content = arguments.get("content", "")
        image_path = arguments.get("image_path")
        layout = arguments.get("layout", "title_and_content")

        try:
            filepath = _add_slide_to_existing(pptx_path, title, content, image_path, layout)
            prs = Presentation(filepath)
            total = len(prs.slides)
            return [types.TextContent(
                type="text",
                text=f"Slide added successfully!\n"
                     f"  File: {filepath}\n"
                     f"  Total slides: {total}",
            )]
        except FileNotFoundError as e:
            return [types.TextContent(type="text", text=str(e))]
        except Exception as e:
            return [types.TextContent(type="text", text=f"Error adding slide: {e}")]

    if name == "export_to_pdf":
        pptx_path = arguments.get("pptx_path")
        if not pptx_path:
            return [types.TextContent(type="text", text="Missing required parameter: pptx_path")]

        output_path = arguments.get("output_path")

        try:
            filepath = _export_to_pdf(pptx_path, output_path)
            return [types.TextContent(
                type="text",
                text=f"PDF exported successfully!\n  File: {filepath}",
            )]
        except FileNotFoundError as e:
            return [types.TextContent(type="text", text=str(e))]
        except RuntimeError as e:
            return [types.TextContent(type="text", text=str(e))]
        except Exception as e:
            return [types.TextContent(type="text", text=f"Error exporting to PDF: {e}")]

    if name == "create_thumbnail":
        pptx_path = arguments.get("pptx_path")
        if not pptx_path:
            return [types.TextContent(type="text", text="Missing required parameter: pptx_path")]

        width = arguments.get("width", 1280)
        output_path = arguments.get("output_path")

        try:
            filepath = _create_thumbnail(pptx_path, width, output_path)
            return [types.TextContent(
                type="text",
                text=f"Thumbnail created successfully!\n"
                     f"  File: {filepath}\n"
                     f"  Width: {width}px",
            )]
        except FileNotFoundError as e:
            return [types.TextContent(type="text", text=str(e))]
        except ValueError as e:
            return [types.TextContent(type="text", text=str(e))]
        except Exception as e:
            return [types.TextContent(type="text", text=f"Error creating thumbnail: {e}")]

    return [types.TextContent(type="text", text=f"Unknown tool: {name}")]


async def main():
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="mcp-presentation",
                server_version="1.0.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


if __name__ == "__main__":
    asyncio.run(main())
