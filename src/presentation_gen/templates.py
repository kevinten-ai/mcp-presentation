"""Built-in slide templates and styling definitions for presentation generation."""

from dataclasses import dataclass, field
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


@dataclass
class TemplateStyle:
    """Defines visual styling for a presentation template."""
    name: str
    # Title styling
    title_font: str = "Calibri"
    title_size: Pt = field(default_factory=lambda: Pt(32))
    title_color: RGBColor = field(default_factory=lambda: RGBColor(0x33, 0x33, 0x33))
    title_bold: bool = True
    # Body styling
    body_font: str = "Calibri"
    body_size: Pt = field(default_factory=lambda: Pt(18))
    body_color: RGBColor = field(default_factory=lambda: RGBColor(0x55, 0x55, 0x55))
    # Background
    bg_color: RGBColor = field(default_factory=lambda: RGBColor(0xFF, 0xFF, 0xFF))
    # Accent / header bar
    accent_color: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x78, 0xD4))
    accent_secondary: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x5A, 0x9E))
    # Header bar (used by corporate template)
    has_header_bar: bool = False
    header_bar_height: Emu = field(default_factory=lambda: Inches(1.0))
    # Gradient background (used by creative template)
    has_gradient_bg: bool = False
    gradient_color_start: RGBColor = field(default_factory=lambda: RGBColor(0x66, 0x7E, 0xEA))
    gradient_color_end: RGBColor = field(default_factory=lambda: RGBColor(0x76, 0x4B, 0xA2))


# ── Template Definitions ─────────────────────────────────────────────────────

MINIMAL = TemplateStyle(
    name="minimal",
    title_font="Calibri Light",
    title_size=Pt(36),
    title_color=RGBColor(0x1A, 0x1A, 0x1A),
    title_bold=False,
    body_font="Calibri",
    body_size=Pt(18),
    body_color=RGBColor(0x4A, 0x4A, 0x4A),
    bg_color=RGBColor(0xFF, 0xFF, 0xFF),
    accent_color=RGBColor(0x33, 0x33, 0x33),
    accent_secondary=RGBColor(0x99, 0x99, 0x99),
)

CORPORATE = TemplateStyle(
    name="corporate",
    title_font="Calibri",
    title_size=Pt(32),
    title_color=RGBColor(0xFF, 0xFF, 0xFF),
    title_bold=True,
    body_font="Calibri",
    body_size=Pt(18),
    body_color=RGBColor(0x33, 0x33, 0x33),
    bg_color=RGBColor(0xFF, 0xFF, 0xFF),
    accent_color=RGBColor(0x00, 0x52, 0x8A),
    accent_secondary=RGBColor(0x00, 0x78, 0xD4),
    has_header_bar=True,
    header_bar_height=Inches(1.2),
)

CREATIVE = TemplateStyle(
    name="creative",
    title_font="Calibri",
    title_size=Pt(40),
    title_color=RGBColor(0xFF, 0xFF, 0xFF),
    title_bold=True,
    body_font="Calibri",
    body_size=Pt(20),
    body_color=RGBColor(0xF0, 0xF0, 0xF0),
    bg_color=RGBColor(0x2D, 0x1B, 0x69),
    accent_color=RGBColor(0xFF, 0x6B, 0x6B),
    accent_secondary=RGBColor(0x4E, 0xC5, 0xC1),
    has_gradient_bg=True,
    gradient_color_start=RGBColor(0x66, 0x7E, 0xEA),
    gradient_color_end=RGBColor(0x76, 0x4B, 0xA2),
)

# ── Registry ──────────────────────────────────────────────────────────────────

TEMPLATES: dict[str, TemplateStyle] = {
    "minimal": MINIMAL,
    "corporate": CORPORATE,
    "creative": CREATIVE,
}


def get_template(name: str) -> TemplateStyle:
    """Get a template by name. Falls back to 'minimal' for unknown names."""
    return TEMPLATES.get(name, MINIMAL)


# ── Layout constants ──────────────────────────────────────────────────────────

SLIDE_WIDTH = Inches(13.333)   # 16:9 widescreen
SLIDE_HEIGHT = Inches(7.5)

LAYOUTS = {
    "title_and_content": "Title at top, bullet-point content below",
    "title_only": "Large centered title",
    "image_full": "Full-bleed image with optional caption",
    "two_column": "Title plus two content columns side by side",
}
