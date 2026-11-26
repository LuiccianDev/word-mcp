"""
Core document operations for Word Document Server.

This module provides functionality for document lifecycle management,
including creation, copying, merging, and conversion.
"""

import os
import platform
import re
import shutil
import subprocess
import tempfile
from re import Pattern
from typing import Any

from docx import Document
from docx.document import Document as DocumentType
from docx.text.paragraph import Paragraph

from word_mcp.core.document_context import atomic_write_target, document_context
from word_mcp.core.styles import (
    DEFAULT_SETTINGS,
    StyleSettings,
    copy_run_formatting,
    ensure_heading_style,
    ensure_table_style,
)
from word_mcp.core.tables import copy_table
from word_mcp.exception import ConfigurationError, DocumentProcessingError
from word_mcp.utils.document_utils import ensure_docx_extension


def core_create_document(
    filename: str,
    title: str | None = None,
    author: str | None = None,
    style_settings: dict[str, Any] | StyleSettings | None = None,
) -> str:
    """Create a new Word document with optional metadata.

    Args:
        filename: Path to the document to create.
        title: Optional title metadata.
        author: Optional author metadata.

    Returns:
        The path to the created document.
    """
    doc: DocumentType = Document()

    if title:
        doc.core_properties.title = title
    if author:
        doc.core_properties.author = author

    # Initialize style settings
    if isinstance(style_settings, dict):
        settings = StyleSettings(**style_settings)
    elif isinstance(style_settings, StyleSettings):
        settings = style_settings
    else:
        settings = DEFAULT_SETTINGS

    ensure_heading_style(doc, settings=settings)
    ensure_table_style(doc)

    with atomic_write_target(filename) as tmp_path:
        doc.save(str(tmp_path))
    return filename


def core_copy_document(
    source_filename: str, destination_filename: str | None = None
) -> str:
    """Create a copy of a Word document atomically, raising on failure."""
    source_filename = ensure_docx_extension(source_filename)
    if not os.path.exists(source_filename):
        raise FileNotFoundError(f"Source document '{source_filename}' does not exist.")

    if destination_filename:
        dest_path = ensure_docx_extension(destination_filename)
    else:
        base, ext = os.path.splitext(source_filename)
        dest_path = f"{base}_copy{ext}"

    with atomic_write_target(dest_path) as tmp_path:
        shutil.copy2(source_filename, str(tmp_path))
    return dest_path


def _copy_paragraph_preserving_runs(
    paragraph: Paragraph, target_doc: DocumentType
) -> None:
    """Copy a paragraph into target_doc while preserving individual run formatting."""
    from docx.oxml.ns import qn
    from docx.text.run import Run

    has_child_runs = any(
        child.tag in (qn("w:r"), qn("w:hyperlink")) for child in paragraph._p
    )
    if not has_child_runs:
        new_paragraph = target_doc.add_paragraph(paragraph.text)
    else:
        new_paragraph = target_doc.add_paragraph()
        for child in paragraph._p:
            if child.tag == qn("w:r"):
                src_run = Run(child, paragraph)
                new_run = new_paragraph.add_run(src_run.text)
                copy_run_formatting(src_run, new_run)
            elif child.tag == qn("w:hyperlink"):
                for h_child in child:
                    if h_child.tag == qn("w:r"):
                        src_run = Run(h_child, paragraph)
                        new_run = new_paragraph.add_run(src_run.text)
                        copy_run_formatting(src_run, new_run)

    try:
        if paragraph.style and paragraph.style.name in target_doc.styles:
            new_paragraph.style = paragraph.style.name
    except (KeyError, AttributeError):
        pass
    if paragraph.paragraph_format.alignment is not None:
        new_paragraph.paragraph_format.alignment = paragraph.paragraph_format.alignment


def core_merge_documents(
    target_filename: str,
    source_filenames: list[str],
    add_page_breaks: bool = True,
) -> str:
    """Merge multiple Word documents into a single document.

    Args:
        target_filename: Path to the target document.
        source_filenames: List of source document paths.
        add_page_breaks: Whether to add page breaks between documents.

    Returns:
        The path to the merged document.
    """
    target_doc: DocumentType = Document()

    for i, filename in enumerate(source_filenames):
        doc_filename = ensure_docx_extension(filename)
        with document_context(doc_filename, mode="read") as source_doc:
            if add_page_breaks and i > 0:
                target_doc.add_page_break()

            para_map = {id(p._p): p for p in source_doc.paragraphs}
            table_map = {id(t._tbl): t for t in source_doc.tables}
            for child in source_doc.element.body:
                cid = id(child)
                if cid in para_map:
                    _copy_paragraph_preserving_runs(para_map[cid], target_doc)
                elif cid in table_map:
                    copy_table(table_map[cid], target_doc)

    with atomic_write_target(target_filename) as tmp_path:
        target_doc.save(str(tmp_path))
    return target_filename


def core_convert_to_pdf(filename: str, output_filename: str | None = None) -> str:
    """Convert a Word document to PDF format."""
    filename = os.path.abspath(filename)
    if not output_filename:
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}.pdf"
    elif not output_filename.lower().endswith(".pdf"):
        output_filename = f"{output_filename}.pdf"

    if not os.path.isabs(output_filename):
        output_filename = os.path.abspath(output_filename)

    output_dir = os.path.dirname(output_filename) or os.path.abspath(".")
    os.makedirs(output_dir, exist_ok=True)

    system = platform.system()

    if system == "Windows":
        try:
            from docx2pdf import convert

            convert(filename, output_filename)
            return output_filename
        except ImportError as exc:
            raise ConfigurationError(
                "docx2pdf is required for PDF conversion on Windows.",
                suggestion="Install docx2pdf and ensure Microsoft Word is available.",
            ) from exc
        except Exception as exc:
            raise DocumentProcessingError(
                f"PDF conversion failed: {exc}",
                suggestion="Verify Microsoft Word is available and the document is not locked.",
            ) from exc

    if system in ["Linux", "Darwin"]:
        lo_commands = (
            ["libreoffice", "soffice"]
            if system == "Linux"
            else ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
        )

        last_error: Exception | None = None
        for cmd_name in lo_commands:
            try:
                with tempfile.TemporaryDirectory(
                    dir=output_dir, prefix=".word_mcp_pdf_"
                ) as staging_dir:
                    cmd = [
                        cmd_name,
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        staging_dir,
                        filename,
                    ]
                    subprocess.run(
                        cmd, capture_output=True, text=True, timeout=60, check=True
                    )

                    base_name = os.path.basename(filename)
                    pdf_base_name = os.path.splitext(base_name)[0] + ".pdf"
                    staged_pdf = os.path.join(staging_dir, pdf_base_name)

                    if not os.path.isfile(staged_pdf):
                        raise subprocess.SubprocessError(
                            "LibreOffice exited cleanly without producing a PDF output file."
                        )

                    os.replace(staged_pdf, output_filename)
                return output_filename
            except (subprocess.SubprocessError, FileNotFoundError) as exc:
                last_error = exc
                continue

        try:
            from docx2pdf import convert

            convert(filename, output_filename)
            return output_filename
        except ImportError as exc:
            raise ConfigurationError(
                f"PDF conversion requires LibreOffice or docx2pdf ({last_error}).",
                suggestion="Install LibreOffice or docx2pdf to enable PDF export.",
            ) from exc
        except Exception as exc:
            raise DocumentProcessingError(
                f"PDF conversion failed. LibreOffice error: {last_error}. Fallback error: {exc}",
                suggestion="Verify LibreOffice or Microsoft Word is installed and accessible.",
            ) from exc

    raise ConfigurationError(
        f"Unsupported platform for PDF conversion: {system}",
        suggestion="Run PDF conversion on Windows, macOS, or Linux.",
    )


def core_get_paragraph_text(doc: DocumentType, paragraph_index: int) -> dict[str, Any]:
    """Get text and metadata from a specific paragraph.

    Args:
        doc: The Word document object.
        paragraph_index: Index of the paragraph.

    Returns:
        Dictionary with paragraph text and metadata.
    """
    if not 0 <= paragraph_index < len(doc.paragraphs):
        raise IndexError(f"Paragraph index {paragraph_index} out of range.")

    paragraph: Paragraph = doc.paragraphs[paragraph_index]

    return {
        "index": paragraph_index,
        "text": paragraph.text,
        "style": paragraph.style.name if paragraph.style else "Normal",
        "is_heading": (
            paragraph.style.name.startswith("Heading") if paragraph.style else False
        ),
    }


def core_find_text(
    doc: DocumentType,
    text_to_find: str,
    match_case: bool = True,
    whole_word: bool = False,
) -> dict[str, Any]:
    """Find all occurrences of specific text in a Word document.

    Args:
        doc: The Word document object.
        text_to_find: Text to search for.
        match_case: If True, performs a case-sensitive search.
        whole_word: If True, matches whole words only.

    Returns:
        A dictionary with search results.
    """
    if not text_to_find:
        raise ValueError("Search text cannot be empty")

    search_pattern = _create_search_pattern(text_to_find, match_case, whole_word)
    all_occurrences: list[dict[str, Any]] = []

    # Search in paragraphs
    for i, para in enumerate(doc.paragraphs):
        location = f"Paragraph {i}"
        all_occurrences.extend(_search_in_element(para, search_pattern, location))

    # Search in tables
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                location = f"Table {t_idx}, Row {r_idx}, Cell {c_idx}"
                all_occurrences.extend(
                    _search_in_element(cell, search_pattern, location)
                )

    return {
        "query": text_to_find,
        "match_case": match_case,
        "whole_word": whole_word,
        "occurrences": all_occurrences,
        "total_count": len(all_occurrences),
    }


def _create_search_pattern(
    text_to_find: str, match_case: bool, whole_word: bool
) -> Pattern[str]:
    """Create a regex pattern for searching text."""
    pattern = re.escape(text_to_find)
    if whole_word:
        pattern = r"\b" + pattern + r"\b"

    flags = 0 if match_case else re.IGNORECASE
    return re.compile(pattern, flags)


def _search_in_element(
    element: Any, pattern: Pattern[str], location_prefix: str
) -> list[dict[str, Any]]:
    """Search for a pattern within a document element (paragraph or cell)."""
    occurrences = []
    for match in pattern.finditer(element.text):
        context = element.text[:100] + ("..." if len(element.text) > 100 else "")
        occurrences.append(
            {
                "location": location_prefix,
                "position": match.start(),
                "match": match.group(0),
                "context": context,
            }
        )
    return occurrences


def _heading_level(paragraph: Paragraph) -> int | None:
    """Return heading level (1-9) if paragraph has a Heading style, else None."""
    style_name = paragraph.style.name if paragraph.style else ""
    if style_name.startswith("Heading "):
        suffix = style_name.split(" ", 1)[1].strip()
        if suffix.isdigit():
            return int(suffix)
    return None


def core_read_range(
    doc: DocumentType,
    start_index: int = 0,
    end_index: int | None = None,
    section_heading: str | None = None,
    max_chars: int = 4000,
) -> dict[str, Any]:
    """Read a context-bounded range of paragraphs by index or heading section."""
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    char_budget = max(1, int(max_chars))
    section_found: bool | None = None

    if section_heading is not None:
        needle = section_heading.strip().lower()
        sec_start: int | None = None
        sec_level: int | None = None
        for idx, para in enumerate(paragraphs):
            lvl = _heading_level(para)
            if lvl is not None and needle in para.text.lower():
                sec_start = idx
                sec_level = lvl
                break

        if sec_start is None:
            return {
                "paragraphs": [],
                "total_paragraphs": total,
                "returned_count": 0,
                "truncated": False,
                "next_index": None,
                "section_found": False,
            }

        section_found = True
        sec_end = total
        for idx in range(sec_start + 1, total):
            lvl = _heading_level(paragraphs[idx])
            if lvl is not None and sec_level is not None and lvl <= sec_level:
                sec_end = idx
                break
        slice_start = max(sec_start, start_index)
        slice_end = sec_end if end_index is None else min(sec_end, end_index + 1)
    else:
        slice_start = max(0, start_index)
        slice_end = total if end_index is None else min(total, end_index + 1)

    collected: list[dict[str, Any]] = []
    used_chars = 0
    truncated = False
    next_index: int | None = None

    for idx in range(slice_start, max(slice_start, slice_end)):
        para = paragraphs[idx]
        text = para.text
        if collected and used_chars + len(text) > char_budget:
            truncated = True
            next_index = idx
            break
        used_chars += len(text)
        collected.append(
            {
                "index": idx,
                "text": text,
                "style": para.style.name if para.style else "Normal",
            }
        )

    result: dict[str, Any] = {
        "paragraphs": collected,
        "total_paragraphs": total,
        "returned_count": len(collected),
        "truncated": truncated,
        "next_index": next_index,
    }
    if section_found is not None:
        result["section_found"] = section_found
    return result
