"""Formatting-related core operations for Word Document Server.

This module provides functions for formatting text, paragraphs, and other
elements within Word documents.
"""

from typing import Any

from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from word_mcp.core.styles import DEFAULT_SETTINGS


def _resolve_color(color: str) -> tuple[RGBColor, str]:
    lower = color.lower()
    if lower in DEFAULT_SETTINGS.colors:
        return DEFAULT_SETTINGS.colors[lower], lower
    try:
        return RGBColor.from_string(color.lstrip("#")), color
    except (ValueError, AttributeError):
        return RGBColor(0, 0, 0), "default (black)"


def _copy_base_props(
    target_run: Any,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    font_name: str | None,
    font_size: Any,
    color_rgb: RGBColor | None,
) -> None:
    if bold is not None:
        target_run.bold = bold
    if italic is not None:
        target_run.italic = italic
    if underline is not None:
        target_run.underline = underline
    if font_name is not None:
        target_run.font.name = font_name
    if font_size is not None:
        target_run.font.size = font_size
    if color_rgb is not None:
        target_run.font.color.rgb = color_rgb


def core_format_text(
    doc: DocumentType,
    paragraph_index: int,
    start_pos: int,
    end_pos: int,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    font_size: int | None = None,
    font_name: str | None = None,
) -> dict[str, Any]:
    """Format a specific range of text within a paragraph without dropping sibling styling."""
    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        raise IndexError(f"Paragraph index {paragraph_index} is out of range.")

    paragraph = doc.paragraphs[paragraph_index]
    text = paragraph.text

    if start_pos < 0 or end_pos > len(text) or start_pos >= end_pos:
        raise ValueError(
            f"Invalid text positions: {start_pos}-{end_pos} for text of length {len(text)}"
        )

    target_text = text[start_pos:end_pos]
    snapshots: list[tuple[str, Any, Any, Any, Any, Any, Any]] = []
    for run in paragraph.runs:
        if not run.text:
            continue
        snapshots.append(
            (
                run.text,
                run.bold,
                run.italic,
                run.underline,
                run.font.name if run.font else None,
                run.font.size if run.font else None,
                run.font.color.rgb if run.font and run.font.color else None,
            )
        )
    if not snapshots and text:
        snapshots = [(text, None, None, None, None, None, None)]

    for r_el in list(paragraph._p.findall(qn("w:r"))):
        paragraph._p.remove(r_el)

    formatting_applied: dict[str, Any] = {}
    if bold is not None:
        formatting_applied["bold"] = bold
    if italic is not None:
        formatting_applied["italic"] = italic
    if underline is not None:
        formatting_applied["underline"] = underline

    resolved_rgb: RGBColor | None = None
    if color:
        resolved_rgb, color_label = _resolve_color(color)
        formatting_applied["color"] = color_label
    if font_size is not None:
        formatting_applied["font_size"] = font_size
    if font_name:
        formatting_applied["font_name"] = font_name
    target_pt = Pt(font_size) if font_size is not None else None

    cursor = 0
    for r_text, r_bold, r_italic, r_underline, r_fname, r_fsize, r_color in snapshots:
        r_end = cursor + len(r_text)
        segments: list[tuple[str, bool]] = []
        if cursor < start_pos:
            prefix_end = min(r_end, start_pos)
            if prefix_end > cursor:
                segments.append((r_text[: prefix_end - cursor], False))
        overlap_start = max(cursor, start_pos)
        overlap_end = min(r_end, end_pos)
        if overlap_start < overlap_end:
            segments.append(
                (r_text[overlap_start - cursor : overlap_end - cursor], True)
            )
        if r_end > end_pos:
            suffix_start = max(cursor, end_pos)
            if r_end > suffix_start:
                segments.append((r_text[suffix_start - cursor :], False))

        for seg_text, in_target in segments:
            new_run = paragraph.add_run(seg_text)
            _copy_base_props(
                new_run, r_bold, r_italic, r_underline, r_fname, r_fsize, r_color
            )
            if in_target:
                _copy_base_props(
                    new_run, bold, italic, underline, font_name, target_pt, resolved_rgb
                )

        cursor = r_end

    return {
        "paragraph_index": paragraph_index,
        "text_formatted": target_text,
        "formatting_applied": formatting_applied,
    }
