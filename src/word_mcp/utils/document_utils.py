"""Document utility functions for Word Document Server.

This module provides pure utility functions for extracting text and structure,
generating atomic document copies, and normalizing .docx extensions.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentType


def core_extract_document_text(doc: DocumentType) -> dict[str, Any]:
    """Extract all non-empty paragraph and table text from an open Document."""
    text_parts: list[str] = []

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            text_parts.append(paragraph.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip():
                        text_parts.append(paragraph.text)

    return {"status": "success", "text": "\n".join(text_parts)}


def core_get_document_structure(doc: DocumentType) -> dict[str, Any]:
    """Extract paragraph and table structure previews from an open Document."""
    structure: dict[str, list[dict[str, Any]]] = {"paragraphs": [], "tables": []}

    for para_idx, paragraph in enumerate(doc.paragraphs):
        if not paragraph.text.strip():
            continue

        preview_text = paragraph.text[:100] + (
            "..." if len(paragraph.text) > 100 else ""
        )
        structure["paragraphs"].append(
            {
                "index": para_idx,
                "text": preview_text,
                "style": paragraph.style.name if paragraph.style else "Normal",
            }
        )

    for table_idx, table in enumerate(doc.tables):
        table_data: dict[str, Any] = {
            "index": table_idx,
            "rows": len(table.rows),
            "columns": len(table.columns),
            "preview": [],
        }

        max_preview_rows = min(3, len(table.rows))
        max_preview_cols = min(3, len(table.columns))

        for row_idx in range(max_preview_rows):
            row_data: list[str] = []
            for col_idx in range(max_preview_cols):
                try:
                    cell_text = table.cell(row_idx, col_idx).text
                    preview_text = cell_text[:20] + (
                        "..." if len(cell_text) > 20 else ""
                    )
                    row_data.append(preview_text)
                except IndexError:
                    row_data.append("N/A")

            if row_data:
                table_data["preview"].append(row_data)
        structure["tables"].append(table_data)

    return structure


def extract_document_text(filename: str) -> dict[str, Any]:
    """Extract all text from a Word document in a single open."""
    doc = Document(filename)
    return core_extract_document_text(doc)


def get_document_structure(filename: str) -> dict[str, Any]:
    """Get the structure of a Word document in a single open."""
    doc = Document(filename)
    return core_get_document_structure(doc)


def create_document_copy(
    source_path: str, dest_path: str | None = None
) -> tuple[bool, str, str | None]:
    """Create a copy of a document atomically via sibling temp replacement."""
    if not os.path.exists(source_path):
        return False, f"Source document '{source_path}' does not exist.", None

    if not dest_path:
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"

    target = Path(dest_path)
    if not target.parent.exists():
        return False, f"Parent directory '{target.parent}' does not exist.", None

    fd, tmp_name = tempfile.mkstemp(
        dir=str(target.parent),
        prefix=f".{target.stem}.",
        suffix=target.suffix or ".tmp",
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        shutil.copy2(source_path, str(tmp_path))
        os.replace(tmp_path, target)
        return True, f"Document successfully copied to '{dest_path}'.", dest_path
    except (OSError, shutil.Error) as exc:
        with contextlib.suppress(OSError):
            tmp_path.unlink(missing_ok=True)
        return False, f"Failed to copy document: {exc}", None


def ensure_docx_extension(filename: str) -> str:
    """Ensure a filename has a .docx extension (case-insensitive check)."""
    if not filename.lower().endswith(".docx"):
        return filename + ".docx"
    return filename
