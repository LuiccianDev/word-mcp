"""Unified document profile registry and typography helpers for Word MCP.

Provides context-adaptive profiles (``modern``, ``formal``, ``report``,
``academic``, ``apa``, ``compact``) and standard page sizes (``letter``, ``a4``,
``legal``, ``a5``) shared across batch, content, and layout operations.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any

from docx.document import Document as DocumentType
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Length, Pt, RGBColor

from word_mcp.core.styles import ensure_heading_style


PAGE_SIZES: dict[str, tuple[Length, Length]] = {
    "letter": (Inches(8.5), Inches(11.0)),
    "a4": (Cm(21.0), Cm(29.7)),
    "legal": (Inches(8.5), Inches(14.0)),
    "a5": (Cm(14.8), Cm(21.0)),
}

_DEFAULT_HEADINGS_SANS: dict[int, dict[str, Any]] = {
    1: {"size": 18.0, "bold": True, "italic": False, "color": (31, 78, 121), "space_before": 14.0, "space_after": 6.0},
    2: {"size": 14.0, "bold": True, "italic": False, "color": (46, 117, 182), "space_before": 12.0, "space_after": 4.0},
    3: {"size": 12.0, "bold": True, "italic": False, "color": (68, 84, 106), "space_before": 10.0, "space_after": 4.0},
}

_DEFAULT_HEADINGS_SERIF: dict[int, dict[str, Any]] = {
    1: {"size": 16.0, "bold": True, "italic": False, "color": (0, 0, 0), "space_before": 12.0, "space_after": 6.0},
    2: {"size": 14.0, "bold": True, "italic": False, "color": (0, 0, 0), "space_before": 10.0, "space_after": 4.0},
    3: {"size": 12.0, "bold": True, "italic": False, "color": (0, 0, 0), "space_before": 8.0, "space_after": 4.0},
}

_APA_HEADINGS: dict[int, dict[str, Any]] = {
    1: {"size": 12.0, "bold": True, "italic": False, "color": (0, 0, 0), "alignment": WD_ALIGN_PARAGRAPH.CENTER, "space_before": 12.0, "space_after": 0.0},
    2: {"size": 12.0, "bold": True, "italic": False, "color": (0, 0, 0), "alignment": WD_ALIGN_PARAGRAPH.LEFT, "space_before": 12.0, "space_after": 0.0},
    3: {"size": 12.0, "bold": True, "italic": True, "color": (0, 0, 0), "alignment": WD_ALIGN_PARAGRAPH.LEFT, "space_before": 12.0, "space_after": 0.0},
    4: {"size": 12.0, "bold": True, "italic": False, "color": (0, 0, 0), "alignment": WD_ALIGN_PARAGRAPH.LEFT, "space_before": 6.0, "space_after": 0.0},
    5: {"size": 12.0, "bold": True, "italic": True, "color": (0, 0, 0), "alignment": WD_ALIGN_PARAGRAPH.LEFT, "space_before": 6.0, "space_after": 0.0},
}

PRESETS: dict[str, dict[str, Any]] = {
    "modern": {
        "name": "modern",
        "font_name": "Calibri",
        "font_size": 11.0,
        "line_spacing": 1.15,
        "space_after": 6.0,
        "page_size": "letter",
        "margins": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
        "headings": _DEFAULT_HEADINGS_SANS,
        "table_header_bg": "1F4E79",
        "hanging_indent_cm": 1.27,
    },
    "formal": {
        "name": "formal",
        "font_name": "Times New Roman",
        "font_size": 12.0,
        "line_spacing": 1.5,
        "space_after": 6.0,
        "page_size": "letter",
        "margins": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5},
        "headings": _DEFAULT_HEADINGS_SERIF,
        "table_header_bg": "2C3E50",
        "hanging_indent_cm": 1.27,
    },
    "report": {
        "name": "report",
        "font_name": "Calibri",
        "font_size": 11.0,
        "line_spacing": 1.15,
        "space_after": 6.0,
        "page_size": "a4",
        "margins": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
        "headings": {
            1: {"size": 18.0, "bold": True, "italic": False, "color": (27, 54, 93), "space_before": 14.0, "space_after": 6.0},
            2: {"size": 14.0, "bold": True, "italic": False, "color": (27, 54, 93), "space_before": 12.0, "space_after": 4.0},
            3: {"size": 12.0, "bold": True, "italic": False, "color": (51, 78, 104), "space_before": 10.0, "space_after": 4.0},
        },
        "table_header_bg": "1B365D",
        "hanging_indent_cm": 1.27,
    },
    "academic": {
        "name": "academic",
        "font_name": "Times New Roman",
        "font_size": 12.0,
        "line_spacing": 2.0,
        "space_after": 0.0,
        "page_size": "letter",
        "margins": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
        "headings": _DEFAULT_HEADINGS_SERIF,
        "table_header_bg": "333333",
        "hanging_indent_cm": 1.27,
    },
    "apa": {
        "name": "apa",
        "font_name": "Times New Roman",
        "font_size": 12.0,
        "line_spacing": 2.0,
        "space_after": 0.0,
        "page_size": "letter",
        "margins": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
        "headings": _APA_HEADINGS,
        "table_header_bg": "1A1A1A",
        "hanging_indent_cm": 1.27,
    },
    "compact": {
        "name": "compact",
        "font_name": "Arial",
        "font_size": 10.0,
        "line_spacing": 1.0,
        "space_after": 4.0,
        "page_size": "letter",
        "margins": {"top": 2.0, "bottom": 2.0, "left": 2.0, "right": 2.0},
        "headings": _DEFAULT_HEADINGS_SANS,
        "table_header_bg": "374151",
        "hanging_indent_cm": 1.0,
    },
}


def resolve_profile(preset: str | None) -> dict[str, Any]:
    """Return a copy of the requested document profile or raise ValueError."""
    key = (preset or "modern").strip().lower()
    if key not in PRESETS:
        valid = ", ".join(sorted(PRESETS))
        raise ValueError(f"Unknown preset '{preset}'. Valid presets: {valid}.")
    return deepcopy(PRESETS[key])


def resolve_page_size(page_size: str) -> tuple[Length, Length]:
    """Return (width, height) dimensions for a standard page size name."""
    key = page_size.strip().lower()
    if key not in PAGE_SIZES:
        valid = ", ".join(sorted(PAGE_SIZES))
        raise ValueError(f"Unknown page_size '{page_size}'. Valid sizes: {valid}.")
    return PAGE_SIZES[key]


def apply_profile_typography(doc: DocumentType, profile: dict[str, Any]) -> None:
    """Apply profile font, spacing, and heading hierarchy across document styles."""
    font_name = str(profile["font_name"])
    font_size = float(profile["font_size"])
    line_spacing = float(profile["line_spacing"])
    space_after = float(profile["space_after"])

    for style_name in ("Normal", "List Bullet", "List Number"):
        try:
            style = doc.styles[style_name]
            style.font.name = font_name
            style.font.size = Pt(font_size)
            style.paragraph_format.line_spacing = line_spacing
            style.paragraph_format.space_after = Pt(space_after)
        except (KeyError, AttributeError):
            continue

    ensure_heading_style(doc)
    headings_cfg: dict[int, dict[str, Any]] = profile.get("headings", {})
    for level in range(1, 10):
        style_name = f"Heading {level}"
        try:
            h_style = doc.styles[style_name]
        except KeyError:
            continue
        cfg = headings_cfg.get(level)
        h_style.font.name = font_name
        h_style.paragraph_format.keep_with_next = True
        if cfg:
            h_style.font.size = Pt(float(cfg["size"]))
            h_style.font.bold = bool(cfg.get("bold", True))
            h_style.font.italic = bool(cfg.get("italic", False))
            if "color" in cfg:
                c = cfg["color"]
                h_style.font.color.rgb = (
                    c if isinstance(c, RGBColor) else RGBColor(*c)
                )
            if "alignment" in cfg:
                h_style.paragraph_format.alignment = cfg["alignment"]
            if "space_before" in cfg:
                h_style.paragraph_format.space_before = Pt(float(cfg["space_before"]))
            if "space_after" in cfg:
                h_style.paragraph_format.space_after = Pt(float(cfg["space_after"]))
        else:
            h_style.font.size = Pt(font_size)
            h_style.font.bold = True

    for special, scale, bold in (("Title", 1.8, True), ("Subtitle", 1.2, False)):
        try:
            s = doc.styles[special]
            s.font.name = font_name
            s.font.size = Pt(round(font_size * scale, 1))
            s.font.bold = bold
        except (KeyError, AttributeError):
            continue
