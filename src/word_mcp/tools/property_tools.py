"""FastMCP tools for document properties and page layout."""

from __future__ import annotations

from typing import Any

from word_mcp.core import (
    core_get_core_properties,
    core_set_core_properties,
)
from word_mcp.core.document_context import document_context
from word_mcp.exception import ValidationError
from word_mcp.tools.format_tools import set_page_layout
from word_mcp.validation.document_validators import (
    validate_docx_read,
    validate_docx_write,
)


@validate_docx_read("filename")
async def get_core_properties(filename: str) -> dict[str, Any]:
    """Get the core metadata properties of the document (author, title, etc.)."""
    with document_context(filename, mode="read") as doc:
        props = core_get_core_properties(doc)

    return {
        "status": "success",
        "properties": props,
        "filename": filename,
    }


@validate_docx_write("filename")
async def set_core_properties(
    filename: str,
    author: str | None = None,
    title: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
    category: str | None = None,
) -> dict[str, Any]:
    """Set the core metadata properties of the document."""
    kwargs: dict[str, str] = {}
    if author is not None:
        kwargs["author"] = author
    if title is not None:
        kwargs["title"] = title
    if subject is not None:
        kwargs["subject"] = subject
    if keywords is not None:
        kwargs["keywords"] = keywords
    if comments is not None:
        kwargs["comments"] = comments
    if category is not None:
        kwargs["category"] = category

    if not kwargs:
        raise ValidationError(
            "No properties provided to update.",
            suggestion="Specify at least one property such as author, title, subject, keywords, comments, or category.",
        )

    with document_context(filename, mode="write") as doc:
        core_set_core_properties(doc, **kwargs)

    return {
        "status": "success",
        "message": f"Successfully updated properties: {list(kwargs.keys())}",
        "filename": filename,
    }


__all__ = [
    "get_core_properties",
    "set_core_properties",
    "set_page_layout",
]
