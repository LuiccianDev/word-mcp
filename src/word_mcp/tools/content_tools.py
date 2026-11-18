"""Content tools for Word Document Server.

This module provides functions to add and manipulate various types of content
in Word documents, including headings, paragraphs, tables, images, and page breaks.
"""

from __future__ import annotations

from typing import Any

from word_mcp.core import (
    core_add_bookmark,
    core_add_heading,
    core_add_hyperlink,
    core_add_page_break,
    core_add_paragraph,
    core_add_picture,
    core_add_table,
    core_add_table_of_contents,
    core_delete_paragraph,
    core_find_and_replace_text,
)
from word_mcp.core.content import core_batch_edit
from word_mcp.core.document_context import document_context
from word_mcp.exception import ValidationError
from word_mcp.models.response_models import (
    HeadingResult,
    OperationResult,
    ParagraphResult,
    PictureResult,
    TableOfContentsResult,
    TableResult,
)
from word_mcp.validation.document_validators import (
    PathAccess,
    PathRule,
    tool_boundary,
    validate_docx_write,
)


def _validate_heading_level(level: int) -> int:
    """Validate and normalize heading level."""
    try:
        level_int = int(level)
        if not 1 <= level_int <= 9:
            raise ValueError(f"Heading level must be between 1 and 9, got {level_int}")
        return level_int
    except (ValueError, TypeError) as error:
        raise ValidationError(
            f"Invalid heading level: {level}. Must be an integer between 1 and 9."
        ) from error


@validate_docx_write("filename")
async def add_heading(filename: str, text: str, level: int = 1) -> dict[str, Any]:
    """Add a heading to a Word document."""
    heading_level = _validate_heading_level(level)
    with document_context(filename, mode="write") as doc:
        core_add_heading(doc, text, level=heading_level)

    return HeadingResult(
        status="success",
        filename=filename,
        heading_text=text,
        heading_level=heading_level,
        message=f"Heading '{text}' (level {heading_level}) added to {filename}",
    ).model_dump()


@validate_docx_write("filename")
async def add_paragraph(
    filename: str, text: str, style: str | None = None
) -> dict[str, Any]:
    """Add a paragraph to a Word document."""
    with document_context(filename, mode="write") as doc:
        result = core_add_paragraph(doc, text, style=style)

    return ParagraphResult(
        status="success",
        filename=filename,
        paragraph_text=text,
        style_applied=result["style_applied"],
        message=f"Paragraph added to {filename}",
    ).model_dump()


@tool_boundary(
    {
        "filename": PathRule(PathAccess.DOCX_WRITE),
        "image_path": PathRule(PathAccess.SANDBOX_FILE_READ),
    },
    operation="add picture",
)
async def add_picture(
    filename: str, image_path: str, width: float | None = None
) -> dict[str, Any]:
    """Add an image to a Word document."""
    if width is not None and width <= 0:
        raise ValidationError(f"Width must be positive, got {width}")

    with document_context(filename, mode="write") as doc:
        result = core_add_picture(doc, image_path, width=width)

    return PictureResult(
        status="success",
        filename=filename,
        image_name=result["image_name"],
        image_path=image_path,
        width_inches=width,
        message=f"Picture '{result['image_name']}' added to {filename}",
    ).model_dump()


@validate_docx_write("filename")
async def add_table(
    filename: str, rows: int, cols: int, data: list[list[Any]] | None = None
) -> dict[str, Any]:
    """Add a table to a Word document."""
    with document_context(filename, mode="write") as doc:
        core_add_table(doc, rows, cols, data)

    return TableResult(
        status="success",
        filename=filename,
        rows=rows,
        columns=cols,
        data_rows=len(data) if data else 0,
        message=f"Table with {rows} rows and {cols} columns added to {filename}",
    ).model_dump()


@validate_docx_write("filename")
async def add_table_of_contents(
    filename: str, title: str = "Table of Contents", max_level: int = 3
) -> dict[str, Any]:
    """Add a table of contents to a Word document."""
    with document_context(filename, mode="write") as doc:
        result = core_add_table_of_contents(doc, title=title, max_level=max_level)

    return TableOfContentsResult(
        status="success",
        filename=filename,
        title=title,
        entry_count=0,
        max_level=max_level,
        message=result["message"],
    ).model_dump()


@validate_docx_write("filename")
async def add_page_break(filename: str) -> dict[str, Any]:
    """Add a page break to the document."""
    with document_context(filename, mode="write") as doc:
        core_add_page_break(doc)

    return OperationResult(
        status="success",
        message=f"Page break added to {filename}",
        details=None,
    ).model_dump()


@validate_docx_write("filename")
async def delete_paragraph(filename: str, paragraph_index: int) -> dict[str, Any]:
    """Delete a paragraph from a document."""
    with document_context(filename, mode="write") as doc:
        core_delete_paragraph(doc, paragraph_index)

    return OperationResult(
        status="success",
        message=f"Paragraph at index {paragraph_index} deleted successfully.",
        details=None,
    ).model_dump()


@validate_docx_write("filename")
async def search_and_replace(
    filename: str, find_text: str, replace_text: str
) -> dict[str, Any]:
    """Search for text and replace all occurrences in a Word document."""
    if not find_text:
        raise ValidationError("Search text cannot be empty")

    with document_context(filename, mode="write") as doc:
        replacement_count = core_find_and_replace_text(doc, find_text, replace_text)

    if replacement_count > 0:
        return OperationResult(
            status="success",
            message=(
                f"Replaced {replacement_count} occurrence(s) of '{find_text}' "
                f"with '{replace_text}' in {filename}"
            ),
            details={"replacements_made": replacement_count},
        ).model_dump()

    raise ValidationError(
        f"No occurrences of '{find_text}' found in {filename}",
        suggestion="Verify the exact search text with find_text_in_document.",
    )


@validate_docx_write("filename")
async def add_hyperlink(
    filename: str,
    paragraph_index: int,
    text: str,
    url: str | None = None,
    bookmark: str | None = None,
) -> dict[str, Any]:
    """Add a hyperlink to a specific paragraph in a document."""
    with document_context(filename, mode="write") as doc:
        core_add_hyperlink(doc, paragraph_index, text, url=url, bookmark=bookmark)

    target = url if url else f"bookmark: {bookmark}"
    return {
        "status": "success",
        "message": f"Hyperlink to {target} added to paragraph {paragraph_index} in {filename}",
        "filename": filename,
    }


@validate_docx_write("filename")
async def add_bookmark(
    filename: str, paragraph_index: int, name: str
) -> dict[str, Any]:
    """Add a bookmark to a specific paragraph to enable internal linking."""
    with document_context(filename, mode="write") as doc:
        core_add_bookmark(doc, paragraph_index, name)

    return {
        "status": "success",
        "message": f"Bookmark '{name}' added to paragraph {paragraph_index} in {filename}",
        "filename": filename,
    }


@validate_docx_write("filename")
async def batch_edit_document(
    filename: str, operations: list[dict[str, Any]]
) -> dict[str, Any]:
    """Apply multiple content edits atomically in a single document pass."""
    with document_context(filename, mode="write") as doc:
        result = core_batch_edit(doc, operations)
    return {
        "status": "success",
        "filename": filename,
        "message": f"Applied {result['applied_count']} batch edit operations to {filename}",
        **result,
    }
