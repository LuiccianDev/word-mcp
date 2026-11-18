"""Document creation and manipulation tools for Word Document Server.

This module provides functions for creating, reading, and manipulating
Word documents with security checks for file system access.
"""

from __future__ import annotations

import os
from typing import Any

from word_mcp.core import (
    core_convert_to_pdf,
    core_copy_document,
    core_create_document,
    core_find_text,
    core_get_core_properties,
    core_get_paragraph_text,
    core_merge_documents,
)
from word_mcp.core.document import core_read_range
from word_mcp.core.document_context import document_context
from word_mcp.core.tables import core_get_tables
from word_mcp.exception import ConfigurationError
from word_mcp.utils.document_utils import (
    core_extract_document_text,
    core_get_document_structure,
)
from word_mcp.utils.security import get_allowed_directories as _get_allowed_directories
from word_mcp.validation.document_validators import (
    PathAccess,
    PathRule,
    tool_boundary,
    validate_docx_create,
    validate_docx_read,
)


@validate_docx_create("filename")
async def create_document(
    filename: str, title: str | None = None, author: str | None = None
) -> dict[str, Any]:
    """Create a new Word document with optional metadata."""
    path = core_create_document(filename, title=title, author=author)
    return {
        "status": "success",
        "message": f"Document {path} created successfully",
    }


@validate_docx_read("filename")
async def get_document_info(
    filename: str, response_format: str = "markdown"
) -> dict[str, Any]:
    """Get information about a Word document."""
    with document_context(filename, mode="read") as doc:
        properties: dict[str, Any] = core_get_core_properties(doc)
        properties["filename"] = os.path.basename(filename)
        properties["path"] = os.path.abspath(filename)
        properties["size_kb"] = round(os.path.getsize(filename) / 1024, 2)

    return {
        "status": "success",
        "properties": properties,
        "response_format": response_format,
    }


@validate_docx_read("filename", operation="get document text")
async def get_document_text(
    filename: str, response_format: str = "markdown"
) -> dict[str, Any]:
    """Extract all text from a Word document."""
    with document_context(filename, mode="read") as doc:
        return core_extract_document_text(doc)


@validate_docx_read("filename", operation="get document outline")
async def get_document_outline(
    filename: str, response_format: str = "markdown"
) -> dict[str, Any]:
    """Get the structure of a Word document."""
    with document_context(filename, mode="read") as doc:
        return core_get_document_structure(doc)


@tool_boundary(
    {"directory": PathRule(PathAccess.SANDBOX_DIR_READ, required=False)},
    operation="list documents",
    primary_path_param="directory",
)
async def list_available_documents(
    directory: str | None = None,
    page: int = 1,
    page_size: int = 20,
    response_format: str = "markdown",
) -> dict[str, Any]:
    """List all .docx files in the allowed directories with pagination support."""
    if page < 1:
        page = 1
    if page_size < 1:
        page_size = 20
    if page_size > 100:
        page_size = 100

    search_directories = (
        [directory] if directory is not None else _get_allowed_directories()
    )
    if not search_directories:
        raise ConfigurationError(
            "No accessible directories found in MCP configuration."
        )

    all_documents = []
    for search_dir in search_directories:
        if not os.path.exists(search_dir):
            continue

        docx_files = [
            f
            for f in os.listdir(search_dir)
            if f.lower().endswith(".docx") and not f.startswith("~$")
        ]
        for file in sorted(docx_files):
            file_path = os.path.join(search_dir, file)
            size_kb = os.path.getsize(file_path) / 1024
            all_documents.append(
                {
                    "name": file,
                    "path": os.path.abspath(file_path),
                    "size_kb": round(size_kb, 2),
                    "source_directory": search_dir,
                }
            )

    total_documents = len(all_documents)
    offset = (page - 1) * page_size
    paginated_documents = all_documents[offset : offset + page_size]
    has_more = offset + page_size < total_documents

    return {
        "status": "success",
        "message": f"Found {total_documents} Word document(s). Showing page {page}.",
        "documents": paginated_documents,
        "pagination": {
            "page": page,
            "page_size": page_size,
            "total": total_documents,
            "has_more": has_more,
        },
        "response_format": response_format,
    }


@tool_boundary(
    {
        "source_filename": PathRule(PathAccess.DOCX_READ),
        "destination_filename": PathRule(
            PathAccess.DOCX_CREATE, required=False, default_ext=".docx"
        ),
    },
    operation="copy document",
    primary_path_param="source_filename",
)
async def copy_document(
    source_filename: str, destination_filename: str | None = None
) -> dict[str, Any]:
    """Create a copy of a Word document."""
    dest_path = core_copy_document(source_filename, destination_filename)
    return {
        "status": "success",
        "message": f"Document successfully copied to '{dest_path}'.",
        "destination_path": dest_path,
    }


@tool_boundary(
    {
        "target_filename": PathRule(PathAccess.DOCX_CREATE, default_ext=".docx"),
        "source_filenames": PathRule(
            PathAccess.DOCX_READ, many=True, default_ext=".docx"
        ),
    },
    operation="merge documents",
    primary_path_param="target_filename",
)
async def merge_documents(
    target_filename: str,
    source_filenames: list[str],
    add_page_breaks: bool = True,
    response_format: str = "markdown",
) -> dict[str, Any]:
    """Merge multiple Word documents into a single document."""
    path = core_merge_documents(target_filename, source_filenames, add_page_breaks)
    return {
        "status": "success",
        "message": f"Successfully merged {len(source_filenames)} documents into {path}",
        "file_path": path,
    }


@validate_docx_read("filename")
async def get_paragraph_text_from_document(
    filename: str, paragraph_index: int
) -> dict[str, Any]:
    """Get text from a specific paragraph in a Word document."""
    with document_context(filename, mode="read") as doc:
        result = core_get_paragraph_text(doc, paragraph_index)

    return {"status": "success", "paragraph": result}


@validate_docx_read("filename")
async def find_text_in_document(
    filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False
) -> dict[str, Any]:
    """Find occurrences of specific text in a Word document."""
    with document_context(filename, mode="read") as doc:
        result = core_find_text(doc, text_to_find, match_case, whole_word)

    return {"status": "success", "search_results": result}


@tool_boundary(
    {
        "filename": PathRule(PathAccess.DOCX_READ),
        "output_filename": PathRule(
            PathAccess.SANDBOX_FILE_WRITE,
            required=False,
            default_ext=".pdf",
            fallback_from="filename",
            fallback_suffix=".pdf",
        ),
    },
    operation="convert to PDF",
)
async def convert_to_pdf(
    filename: str, output_filename: str | None = None
) -> dict[str, Any]:
    """Convert a Word document to PDF format."""
    pdf_path = core_convert_to_pdf(filename, output_filename)
    return {
        "status": "success",
        "message": "Document successfully converted to PDF",
        "pdf_path": pdf_path,
    }


@validate_docx_read("filename")
async def read_document_range(
    filename: str,
    start_index: int = 0,
    end_index: int | None = None,
    section_heading: str | None = None,
    max_chars: int = 4000,
) -> dict[str, Any]:
    """Read a context-bounded range of paragraphs by index or heading section."""
    with document_context(filename, mode="read") as doc:
        result = core_read_range(
            doc,
            start_index=start_index,
            end_index=end_index,
            section_heading=section_heading,
            max_chars=max_chars,
        )
    return {"status": "success", "filename": filename, **result}


@validate_docx_read("filename")
async def get_document_tables(
    filename: str,
    table_index: int | None = None,
    include_rows: bool = False,
    row_offset: int = 0,
    row_limit: int = 20,
) -> dict[str, Any]:
    """Inspect table metadata and optionally paginate rows without flooding context."""
    with document_context(filename, mode="read") as doc:
        result = core_get_tables(
            doc,
            table_index=table_index,
            include_rows=include_rows,
            row_offset=row_offset,
            row_limit=row_limit,
        )
    return {"status": "success", "filename": filename, **result}
