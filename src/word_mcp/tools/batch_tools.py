"""Batch and composite document building tools for Word MCP Server.

This module provides high-efficiency tools to build full documents or insert
multiple content blocks (headings, paragraphs, tables, lists, breaks) in a single pass.
"""

from __future__ import annotations

from typing import Any

from word_mcp.core.batch import (
    core_build_document,
    core_create_from_markdown,
    render_elements_batch,
    validate_batch_elements,
)
from word_mcp.core.document_context import document_context
from word_mcp.models.response_models import BuildDocumentResult
from word_mcp.validation.document_validators import (
    PathAccess,
    PathRule,
    tool_boundary,
)


@tool_boundary(
    {
        "filename": PathRule(
            PathAccess.DOCX_CREATE, aliases=("file_path",), default_ext=".docx"
        )
    },
    operation="build document",
)
async def build_document(
    filename: str | None = None,
    file_path: str | None = None,
    title: str | None = None,
    author: str | None = None,
    preset: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    orientation: str | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    elements: list[dict[str, Any]] | None = None,
    page_size: str | None = None,
) -> dict[str, Any]:
    """Build a complete Word document with typography, margins, header/footer and content in a single pass."""
    assert filename is not None
    validate_batch_elements(elements)
    target = filename
    with document_context(target, mode="write") as doc:
        res = core_build_document(
            doc=doc,
            title=title,
            author=author,
            preset=preset,
            font_name=font_name,
            font_size=font_size,
            line_spacing=line_spacing,
            space_after=space_after,
            margins=margins,
            orientation=orientation,
            header_text=header_text,
            footer_text=footer_text,
            page_numbers=page_numbers,
            elements=elements,
            page_size=page_size,
        )

    return BuildDocumentResult(
        status="success",
        filename=target,
        total_elements=res["total_elements"],
        element_counts=res["element_counts"],
        summary=res["summary"],
        message=f"Document '{target}' successfully built in a single pass ({res['summary']}).",
    ).model_dump()


@tool_boundary(
    {
        "filename": PathRule(
            PathAccess.DOCX_CREATE, aliases=("file_path",), default_ext=".docx"
        )
    },
    operation="create from markdown",
)
async def create_from_markdown(
    markdown_content: str,
    filename: str | None = None,
    file_path: str | None = None,
    title: str | None = None,
    author: str | None = None,
    preset: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    page_size: str | None = None,
) -> dict[str, Any]:
    """Create a styled Word document directly from Markdown text in a single pass."""
    assert filename is not None
    target = filename
    with document_context(target, mode="write") as doc:
        res = core_create_from_markdown(
            doc=doc,
            markdown_content=markdown_content,
            title=title,
            author=author,
            preset=preset,
            font_name=font_name,
            font_size=font_size,
            line_spacing=line_spacing,
            space_after=space_after,
            margins=margins,
            header_text=header_text,
            footer_text=footer_text,
            page_numbers=page_numbers,
            page_size=page_size,
        )

    return BuildDocumentResult(
        status="success",
        filename=target,
        total_elements=res["total_elements"],
        element_counts=res["element_counts"],
        summary=res["summary"],
        message=f"Document '{target}' created from Markdown in a single pass ({res['summary']}).",
    ).model_dump()


@tool_boundary(
    {"filename": PathRule(PathAccess.DOCX_WRITE, aliases=("file_path",))},
    operation="batch add content",
)
async def batch_add_content(
    elements: list[dict[str, Any]],
    filename: str | None = None,
    file_path: str | None = None,
) -> dict[str, Any]:
    """Append multiple content blocks (headings, paragraphs, tables, lists) to an existing document in one pass."""
    assert filename is not None
    validate_batch_elements(elements)
    target = filename
    with document_context(target, mode="write") as doc:
        counts, total, rendered_summary = render_elements_batch(doc, elements)
    summary = rendered_summary if rendered_summary else "No elements added"

    return BuildDocumentResult(
        status="success",
        filename=target,
        total_elements=total,
        element_counts=counts,
        summary=summary,
        message=f"Appended {total} elements to '{target}' in a single pass.",
    ).model_dump()
