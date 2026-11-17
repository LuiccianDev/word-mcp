"""Batch and composite document building operations for Word MCP.

This module provides high-efficiency in-memory document building, allowing
full documents with typography, margins, line spacing, headers, footers,
tables, lists, and headings to be generated in a single pass.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Any

from docx.document import Document as DocumentType
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.text.run import Run

from word_mcp.core.content import core_add_table_of_contents
from word_mcp.core.presets import (
    apply_profile_typography,
    resolve_page_size,
    resolve_profile,
)
from word_mcp.utils.security import is_path_in_allowed_directories


ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

SUPPORTED_ELEMENT_TYPES = (
    "heading",
    "paragraph",
    "table",
    "bullet_list",
    "numbered_list",
    "page_break",
    "callout",
    "image",
    "toc",
    "references",
    "title_block",
)


@dataclass(frozen=True)
class CalloutStyle:
    """Visual configuration for a styled alert callout box."""

    label: str
    accent_hex: str
    fill_hex: str
    title_hex: str


CALLOUT_VARIANTS: dict[str, CalloutStyle] = {
    "note": CalloutStyle("Note", "0969DA", "F0F6FC", "0550AE"),
    "tip": CalloutStyle("Tip", "1A7F37", "F0FDF4", "116329"),
    "important": CalloutStyle("Important", "8250DF", "F5F0FF", "6639BA"),
    "warning": CalloutStyle("Warning", "9A6700", "FFF8C5", "7D4E00"),
    "caution": CalloutStyle("Caution", "CF222E", "FFEBE9", "A40E26"),
}

_ALERT_MARKER_RE = re.compile(
    r"^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*(.*)$",
    re.IGNORECASE,
)


def _render_callout_box(
    doc: DocumentType,
    text: str,
    style: CalloutStyle,
    title: str | None = None,
) -> None:
    table = doc.add_table(rows=1, cols=1)
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is not None:
        tbl_pr.remove(tbl_w)
    tbl_pr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>'))

    cell = table.cell(0, 0)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_pr.append(
        parse_xml(
            f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{style.fill_hex}"/>'
        )
    )
    tc_pr.append(
        parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{style.accent_hex}"/>'
            f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f"</w:tcBorders>"
        )
    )
    tc_pr.append(
        parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'  <w:top w:w="120" w:type="dxa"/>'
            f'  <w:left w:w="200" w:type="dxa"/>'
            f'  <w:bottom w:w="120" w:type="dxa"/>'
            f'  <w:right w:w="160" w:type="dxa"/>'
            f"</w:tcMar>"
        )
    )

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)

    header_label = str(title).strip() if title else style.label.upper()
    title_run = p.add_run(header_label)
    title_run.bold = True
    title_run.font.color.rgb = RGBColor.from_string(style.title_hex)

    if text:
        title_run.add_break()
        add_inline_formatted_text(p, text)


def apply_page_number_field(paragraph: Any) -> None:
    """Add a native Word PAGE field to a paragraph."""
    run = paragraph.add_run(" ")
    fld_simple = OxmlElement("w:fldSimple")
    fld_simple.set(qn("w:instr"), "PAGE")
    run._r.append(fld_simple)


def apply_document_settings(
    doc: DocumentType,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    orientation: str | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    preset: str | None = None,
    page_size: str | None = None,
) -> dict[str, Any]:
    """Configure document-wide styles, page setup, header, and footer."""
    config = resolve_profile(preset)

    if font_name is not None:
        config["font_name"] = font_name
    if font_size is not None:
        config["font_size"] = font_size
    if line_spacing is not None:
        config["line_spacing"] = line_spacing
    if space_after is not None:
        config["space_after"] = space_after
    if margins is not None:
        config["margins"] = margins
    if page_size is not None:
        config["page_size"] = page_size

    if orientation is not None and orientation.lower() not in ("portrait", "landscape"):
        raise ValueError("Orientation must be 'portrait' or 'landscape'")

    width, height = resolve_page_size(str(config.get("page_size", "letter")))
    apply_profile_typography(doc, config)

    # Configure sections (page size, margins, orientation, headers/footers)
    for section in doc.sections:
        orient = (orientation or "portrait").lower()
        if orient == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = max(width, height)
            section.page_height = min(width, height)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = min(width, height)
            section.page_height = max(width, height)

        m = config.get("margins", {})
        if "top" in m:
            section.top_margin = Cm(float(m["top"]))
        if "bottom" in m:
            section.bottom_margin = Cm(float(m["bottom"]))
        if "left" in m:
            section.left_margin = Cm(float(m["left"]))
        if "right" in m:
            section.right_margin = Cm(float(m["right"]))

        if header_text:
            header_para = section.header.paragraphs[0]
            header_para.text = header_text
            header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if len(header_para.runs) > 0:
                header_para.runs[0].font.size = Pt(8.5)
                header_para.runs[0].font.color.rgb = RGBColor(128, 128, 128)

        if footer_text or page_numbers:
            footer_para = section.footer.paragraphs[0]
            footer_para.text = footer_text or ""
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if page_numbers:
                if footer_text:
                    footer_para.add_run(" | Pág. ")
                apply_page_number_field(footer_para)
            if len(footer_para.runs) > 0:
                footer_para.runs[0].font.size = Pt(8.5)
                footer_para.runs[0].font.color.rgb = RGBColor(128, 128, 128)

    return config


_INLINE_LINK_RE = re.compile(r"^\[([^\[\]]+)\]\(([^\s()]+)\)$")


def _append_inline_hyperlink(paragraph: Any, label: str, target: str) -> bool:
    """Append an external or bookmark w:hyperlink to paragraph; return False if invalid."""
    hyperlink = OxmlElement("w:hyperlink")
    if target.startswith("#") and len(target) > 1:
        hyperlink.set(qn("w:anchor"), target[1:].replace(" ", "_"))
    elif target.startswith(("http://", "https://", "mailto:")):
        r_id = paragraph.part.relate_to(target, RT.HYPERLINK, is_external=True)
        hyperlink.set(qn("r:id"), r_id)
    else:
        return False

    run = Run(OxmlElement("w:r"), paragraph)
    run.text = label
    run.font.color.theme_color = 10
    run.font.underline = True
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)
    return True


def add_inline_formatted_text(paragraph: Any, text: str) -> None:
    """Parse inline [label](url), **bold**, and *italic* markdown in paragraph text."""
    tokens = re.split(r"(\[[^\[\]]+\]\([^\s()]+\)|\*\*.*?\*\*|\*.*?\*)", text)
    for token in tokens:
        if not token:
            continue
        link_match = _INLINE_LINK_RE.match(token)
        if link_match and _append_inline_hyperlink(
            paragraph, link_match.group(1), link_match.group(2)
        ):
            continue
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        else:
            paragraph.add_run(token)


def _render_list_items(
    doc: DocumentType, items: Any, style_name: str, base_level: int = 0
) -> None:
    """Render flat or nested list items with dynamic indentation."""
    if isinstance(items, str):
        items = [items]
    for item in items:
        if isinstance(item, dict):
            text = str(item.get("text", ""))
            level = int(item.get("level", base_level))
            p = doc.add_paragraph(style=style_name)
            if level > 0:
                p.paragraph_format.left_indent = Cm(0.63 * level)
            add_inline_formatted_text(p, text)
            children = item.get("children")
            if children:
                _render_list_items(doc, children, style_name, base_level=level + 1)
        elif isinstance(item, list):
            _render_list_items(doc, item, style_name, base_level=base_level + 1)
        else:
            p = doc.add_paragraph(style=style_name)
            if base_level > 0:
                p.paragraph_format.left_indent = Cm(0.63 * base_level)
            add_inline_formatted_text(p, str(item))


def _format_apa_reference(entry: str | dict[str, Any]) -> str:
    """Format a plain string or structured citation dictionary in APA 7 style."""
    if isinstance(entry, str):
        return entry
    parts: list[str] = []
    authors = str(entry.get("authors", "")).strip()
    year = str(entry.get("year", "")).strip()
    title = str(entry.get("title", "")).strip()
    source = str(entry.get("source", "")).strip()
    doi = str(entry.get("doi") or entry.get("url") or "").strip()

    if authors:
        parts.append(authors if authors.endswith(".") else f"{authors}.")
    if year:
        parts.append(f"({year}).")
    if title:
        parts.append(f"*{title}*." if not title.endswith(".") else f"*{title}*")
    if source:
        parts.append(source if source.endswith(".") else f"{source}.")
    if doi:
        parts.append(doi)
    return " ".join(parts)


def render_element(doc: DocumentType, elem: dict[str, Any]) -> str:
    """Render a single content element dictionary into the docx Document.

    Returns:
        Type of rendered element.
    """
    elem_type = elem.get("type", "paragraph").lower()

    if elem_type == "heading":
        level = int(elem.get("level", 1))
        text = elem.get("text", "")
        p = doc.add_heading(text, level=min(max(level, 1), 9))
        align = elem.get("alignment")
        if align in ALIGNMENT_MAP:
            p.alignment = ALIGNMENT_MAP[align]
        return "heading"

    elif elem_type == "paragraph":
        text = elem.get("text", "")
        p = doc.add_paragraph()
        add_inline_formatted_text(p, text)

        align = elem.get("alignment")
        if align in ALIGNMENT_MAP:
            p.alignment = ALIGNMENT_MAP[align]

        if "space_after" in elem:
            p.paragraph_format.space_after = Pt(float(elem["space_after"]))
        if "line_spacing" in elem:
            p.paragraph_format.line_spacing = float(elem["line_spacing"])
        if elem.get("bold"):
            for run in p.runs:
                run.bold = True
        if elem.get("italic"):
            for run in p.runs:
                run.italic = True
        return "paragraph"

    elif elem_type == "table":
        headers = elem.get("headers", [])
        rows_data = elem.get("rows", [])
        style = elem.get("style", "Table Grid")

        total_rows = len(rows_data) + (1 if headers else 0)
        total_cols = (
            len(headers) if headers else (len(rows_data[0]) if rows_data else 1)
        )

        table = doc.add_table(rows=total_rows, cols=total_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        try:
            table.style = style
        except Exception:
            pass

        current_row = 0
        if headers:
            hdr_cells = table.rows[0].cells
            for idx, h_text in enumerate(headers[:total_cols]):
                hdr_cells[idx].text = str(h_text)
                for run in hdr_cells[idx].paragraphs[0].runs:
                    run.bold = True
            current_row = 1

        for r_data in rows_data:
            if current_row >= len(table.rows):
                break
            row_cells = table.rows[current_row].cells
            for c_idx, cell_value in enumerate(r_data[:total_cols]):
                row_cells[c_idx].text = str(cell_value)
            current_row += 1

        # Add empty spacing paragraph after table
        space_p = doc.add_paragraph()
        space_p.paragraph_format.space_after = Pt(4)
        return "table"

    elif elem_type in ("bullet_list", "bullet", "list"):
        _render_list_items(doc, elem.get("items", []), "List Bullet")
        return "bullet_list"

    elif elem_type in ("numbered_list", "number_list"):
        _render_list_items(doc, elem.get("items", []), "List Number")
        return "numbered_list"

    elif elem_type == "page_break":
        doc.add_page_break()
        return "page_break"

    elif elem_type in ("callout", "quote"):
        text = str(elem.get("text", ""))
        title = elem.get("title")
        raw_variant = elem.get("variant") or elem.get("alert_type")
        if raw_variant is not None:
            variant_key = str(raw_variant).strip().lower()
            if variant_key not in CALLOUT_VARIANTS:
                valid = ", ".join(sorted(CALLOUT_VARIANTS))
                raise ValueError(
                    f"Unsupported callout variant '{raw_variant}'. Valid variants: {valid}."
                )
            _render_callout_box(
                doc,
                text=text,
                style=CALLOUT_VARIANTS[variant_key],
                title=title,
            )
            return "callout"

        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.right_indent = Cm(1.0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        if title:
            title_run = p.add_run(f"{title}: ")
            title_run.bold = True
        add_inline_formatted_text(p, text)
        for run in p.runs:
            run.italic = True
        return "callout"

    elif elem_type in ("image", "picture"):
        raw_path = elem.get("path") or elem.get("image_path")
        if not raw_path or not isinstance(raw_path, str):
            raise ValueError("Image element requires a non-empty 'path' string.")
        abs_path = os.path.abspath(raw_path)
        allowed, reason = is_path_in_allowed_directories(abs_path)
        if not allowed:
            raise ValueError(f"Image path not allowed: {reason}")
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Image file not found: {abs_path}")

        width = elem.get("width")
        if width is not None and float(width) <= 0:
            raise ValueError(f"Image width must be positive, got {width}")

        p = doc.add_paragraph()
        align_str = str(elem.get("alignment", "center")).lower()
        if align_str in ALIGNMENT_MAP:
            p.alignment = ALIGNMENT_MAP[align_str]
        run = p.add_run()
        if width is not None:
            run.add_picture(abs_path, width=Inches(float(width)))
        else:
            run.add_picture(abs_path)

        caption = elem.get("caption")
        if caption:
            cap_p = doc.add_paragraph()
            if align_str in ALIGNMENT_MAP:
                cap_p.alignment = ALIGNMENT_MAP[align_str]
            cap_run = cap_p.add_run(str(caption))
            cap_run.italic = True
            cap_run.font.size = Pt(9.5)
        return "image"

    elif elem_type == "toc":
        core_add_table_of_contents(
            doc,
            title=str(elem.get("title", "Table of Contents")),
            max_level=int(elem.get("max_level", 3)),
            page_break=bool(elem.get("page_break", True)),
        )
        return "toc"

    elif elem_type in ("references", "bibliography"):
        if elem.get("include_heading", True):
            doc.add_heading(
                str(elem.get("title", "References")),
                level=min(max(int(elem.get("level", 1)), 1), 9),
            )
        hanging_cm = float(elem.get("hanging_indent_cm", 1.27))
        for entry in elem.get("items", []):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(hanging_cm)
            p.paragraph_format.first_line_indent = Cm(-hanging_cm)
            add_inline_formatted_text(p, _format_apa_reference(entry))
        return "references"

    elif elem_type == "title_block":
        title_text = elem.get("title")
        if title_text:
            p_title = doc.add_paragraph(str(title_text), style="Title")
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle_text = elem.get("subtitle")
        if subtitle_text:
            p_sub = doc.add_paragraph(str(subtitle_text), style="Subtitle")
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta_parts = [
            str(elem[k]) for k in ("author", "institution", "date") if elem.get(k)
        ]
        if meta_parts:
            p_meta = doc.add_paragraph(" | ".join(meta_parts))
            p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return "title_block"

    valid = ", ".join(SUPPORTED_ELEMENT_TYPES)
    raise ValueError(
        f"Unknown element type '{elem_type}'. Supported element types: {valid}."
    )


def core_build_document(
    doc: DocumentType,
    title: str | None = None,
    author: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    orientation: str | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    preset: str | None = None,
    elements: list[dict[str, Any]] | None = None,
    page_size: str | None = None,
) -> dict[str, Any]:
    """Build a complete Word document in memory in a single pass.

    Returns:
        Summary dictionary with counts of elements created.
    """
    # Core properties
    if title:
        doc.core_properties.title = title
    if author:
        doc.core_properties.author = author

    # Apply typography, margins, header, and footer
    config_applied = apply_document_settings(
        doc=doc,
        font_name=font_name,
        font_size=font_size,
        line_spacing=line_spacing,
        space_after=space_after,
        margins=margins,
        orientation=orientation,
        header_text=header_text,
        footer_text=footer_text,
        page_numbers=page_numbers,
        preset=preset,
        page_size=page_size,
    )

    counts: dict[str, int] = {}
    if elements:
        for elem in elements:
            if isinstance(elem, dict):
                t = render_element(doc, elem)
                counts[t] = counts.get(t, 0) + 1

    total_created = sum(counts.values())
    summary_parts = [f"{v} {k}s" for k, v in counts.items()]
    summary = (
        ", ".join(summary_parts) if summary_parts else "Document layout configured"
    )

    return {
        "status": "success",
        "total_elements": total_created,
        "element_counts": counts,
        "config_applied": config_applied,
        "summary": summary,
    }


def parse_markdown_to_elements(markdown_text: str) -> list[dict[str, Any]]:
    """Convert standard markdown text into a list of structured document elements."""
    elements: list[dict[str, Any]] = []
    lines = markdown_text.splitlines()
    in_table = False
    table_headers: list[str] = []
    table_rows: list[list[str]] = []
    quote_lines: list[str] = []

    def flush_table() -> None:
        nonlocal in_table, table_headers, table_rows
        if in_table and table_headers:
            elements.append(
                {
                    "type": "table",
                    "headers": table_headers,
                    "rows": table_rows,
                    "style": "Table Grid",
                }
            )
        in_table = False
        table_headers = []
        table_rows = []

    def flush_quote() -> None:
        nonlocal quote_lines
        if not quote_lines:
            return
        first_line = quote_lines[0]
        alert_match = _ALERT_MARKER_RE.match(first_line)
        if alert_match:
            variant = alert_match.group(1).lower()
            parts = [alert_match.group(2).strip()] if alert_match.group(2).strip() else []
            parts.extend(q for q in quote_lines[1:] if q)
            elements.append(
                {"type": "callout", "variant": variant, "text": " ".join(parts)}
            )
        else:
            body = " ".join(q for q in quote_lines if q)
            elements.append({"type": "callout", "text": body})
        quote_lines = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            flush_table()
            flush_quote()
            continue

        if line == ">" or line.startswith("> "):
            flush_table()
            quote_lines.append(line[2:].strip() if line.startswith("> ") else "")
            continue
        else:
            flush_quote()

        # Check for table line
        if line.startswith("|") and line.endswith("|"):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            # Check if divider line like |---|---|
            if all(re.match(r"^:?-+:?$", c) for c in cells):
                continue
            if not in_table:
                in_table = True
                table_headers = cells
            else:
                table_rows.append(cells)
            continue
        else:
            flush_table()

        # Page break
        if line in ("---", "***", "___"):
            elements.append({"type": "page_break"})
            continue

        # Headings
        if line.startswith("#"):
            match = re.match(r"^(#{1,6})\s+(.*)$", line)
            if match:
                level = len(match.group(1))
                heading_text = match.group(2).strip()
                elements.append(
                    {"type": "heading", "level": level, "text": heading_text}
                )
                continue

        # Bullet list (supports nested indentation and groups consecutive items)
        if line.startswith(("- ", "* ", "+ ")):
            indent = len(raw_line) - len(raw_line.lstrip(" \t"))
            level = indent // 2
            item_val: str | dict[str, Any] = (
                {"text": line[2:].strip(), "level": level}
                if level > 0
                else line[2:].strip()
            )
            if elements and elements[-1]["type"] == "bullet_list":
                elements[-1]["items"].append(item_val)
            else:
                elements.append({"type": "bullet_list", "items": [item_val]})
            continue

        # Numbered list (supports nested indentation and groups consecutive items)
        num_match = re.match(r"^\d+\.\s+(.*)$", line)
        if num_match:
            indent = len(raw_line) - len(raw_line.lstrip(" \t"))
            level = indent // 2
            item_val = (
                {"text": num_match.group(1).strip(), "level": level}
                if level > 0
                else num_match.group(1).strip()
            )
            if elements and elements[-1]["type"] == "numbered_list":
                elements[-1]["items"].append(item_val)
            else:
                elements.append({"type": "numbered_list", "items": [item_val]})
            continue

        # Standard paragraph
        elements.append({"type": "paragraph", "text": line})

    flush_table()
    flush_quote()
    return elements


def core_create_from_markdown(
    doc: DocumentType,
    markdown_content: str,
    title: str | None = None,
    author: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    preset: str | None = None,
    page_size: str | None = None,
) -> dict[str, Any]:
    """Parse Markdown content and build Word document in a single pass."""
    elements = parse_markdown_to_elements(markdown_content)
    return core_build_document(
        doc=doc,
        title=title,
        author=author,
        font_name=font_name,
        font_size=font_size,
        line_spacing=line_spacing,
        space_after=space_after,
        margins=margins,
        header_text=header_text,
        footer_text=footer_text,
        page_numbers=page_numbers,
        preset=preset,
        elements=elements,
        page_size=page_size,
    )
