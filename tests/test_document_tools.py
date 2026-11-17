"""Tests for the document_tools module."""

import os
from collections.abc import Generator
from pathlib import Path

import pytest
from docx import Document

# Assuming document_tools.py is in src/word_mcp/tools
from word_mcp.tools import document_tools


@pytest.fixture  # type: ignore[misc]
def temp_docx_file(tmp_path: Path) -> Generator[str, None, None]:
    """Create a temporary .docx file for testing."""
    file_path = tmp_path / "test_document.docx"
    doc = Document()
    doc.add_paragraph("This is a test document.")
    doc.save(file_path)
    # Set allowed directory for testing
    os.environ["MCP_ALLOWED_DIRECTORIES"] = str(tmp_path)
    yield str(file_path)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_create_document(tmp_path: Path) -> None:
    """Test creating a new document."""
    os.environ["MCP_ALLOWED_DIRECTORIES"] = str(tmp_path)
    file_path = tmp_path / "new_document.docx"
    result = await document_tools.create_document(
        str(file_path), title="Test Title", author="Test Author"
    )
    assert result["status"] == "success"
    assert isinstance(result, dict)
    assert os.path.exists(file_path)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_get_document_info(temp_docx_file: str) -> None:
    """Test getting document information."""
    result = await document_tools.get_document_info(temp_docx_file)
    assert isinstance(result, dict)
    assert "title" in result["properties"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_get_document_text(temp_docx_file: str) -> None:
    """Test getting document text."""
    result = await document_tools.get_document_text(temp_docx_file)
    assert "This is a test" in result["text"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_get_document_outline(temp_docx_file: str) -> None:
    """Test getting document outline."""
    result = await document_tools.get_document_outline(temp_docx_file)
    assert isinstance(result, dict)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_list_available_documents(temp_docx_file: str, tmp_path: Path) -> None:
    """Test listing available documents."""
    result = await document_tools.list_available_documents()
    assert result["status"] == "success"
    assert len(result["documents"]) > 0


@pytest.mark.asyncio  # type: ignore[misc]
async def test_copy_document(temp_docx_file: str, tmp_path: Path) -> None:
    """Test copying a document."""
    dest_path = tmp_path / "copied_document.docx"
    result = await document_tools.copy_document(temp_docx_file, str(dest_path))
    assert result["status"] == "success"
    assert os.path.exists(dest_path)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_document_range_by_section_and_empty(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test reading a section by heading and handling missing sections or empty docs."""
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "sections.docx"
    doc = Document()
    doc.add_heading("Intro", level=1)
    doc.add_paragraph("First intro paragraph.")
    doc.add_paragraph("Second intro paragraph.")
    doc.add_heading("Details", level=1)
    doc.add_paragraph("Detail paragraph.")
    doc.save(doc_path)

    res = await document_tools.read_document_range(
        str(doc_path), section_heading="Intro"
    )
    assert res["status"] == "success"
    assert res["section_found"] is True
    assert [p["text"] for p in res["paragraphs"]] == [
        "Intro",
        "First intro paragraph.",
        "Second intro paragraph.",
    ]

    paginated = await document_tools.read_document_range(
        str(doc_path), section_heading="Intro", start_index=2
    )
    assert paginated["status"] == "success"
    assert [p["text"] for p in paginated["paragraphs"]] == ["Second intro paragraph."]

    missing = await document_tools.read_document_range(
        str(doc_path), section_heading="NonExistent"
    )
    assert missing["status"] == "success"
    assert missing["section_found"] is False
    assert missing["paragraphs"] == []

    empty_path = tmp_path / "empty.docx"
    Document().save(empty_path)
    empty_res = await document_tools.read_document_range(str(empty_path))
    assert empty_res["status"] == "success"
    assert empty_res["paragraphs"] == []
    assert empty_res["truncated"] is False


@pytest.mark.asyncio  # type: ignore[misc]
async def test_read_document_range_truncates_on_max_chars(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test max_chars cap sets truncated=True and next_index."""
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "large.docx"
    doc = Document()
    for i in range(10):
        doc.add_paragraph(f"Para {i}: " + ("x" * 90))
    doc.save(doc_path)

    res = await document_tools.read_document_range(
        str(doc_path), start_index=0, max_chars=220
    )
    assert res["status"] == "success"
    assert res["truncated"] is True
    assert len(res["paragraphs"]) >= 1
    assert res["next_index"] == res["paragraphs"][-1]["index"] + 1
    assert res["total_paragraphs"] == 10


@pytest.mark.asyncio  # type: ignore[misc]
async def test_get_document_tables_metadata_and_paginated_rows(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test table metadata inspection and row pagination."""
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "tables.docx"
    doc = Document()
    table = doc.add_table(rows=5, cols=2)
    table.rows[0].cells[0].text = "Name"
    table.rows[0].cells[1].text = "Score"
    for r in range(1, 5):
        table.rows[r].cells[0].text = f"Item {r}"
        table.rows[r].cells[1].text = str(r * 10)
    doc.save(doc_path)

    meta = await document_tools.get_document_tables(str(doc_path))
    assert meta["status"] == "success"
    assert meta["total_tables"] == 1
    assert meta["tables"][0]["row_count"] == 5
    assert meta["tables"][0]["col_count"] == 2
    assert meta["tables"][0]["headers"] == ["Name", "Score"]
    assert "rows" not in meta["tables"][0]

    page = await document_tools.get_document_tables(
        str(doc_path),
        table_index=0,
        include_rows=True,
        row_offset=1,
        row_limit=2,
    )
    assert page["status"] == "success"
    assert page["tables"][0]["rows"] == [["Item 1", "10"], ["Item 2", "20"]]
    assert page["tables"][0]["truncated"] is True

    bad_idx = await document_tools.get_document_tables(str(doc_path), table_index=99)
    assert bad_idx["status"] == "error"
