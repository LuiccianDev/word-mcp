"""Tests for actionable hint guidance in document validators."""

from pathlib import Path

from word_mcp.validation.document_validators import (
    _validate_docx_path,
    validate_docx_read,
    validate_file_write,
)


def test_validate_docx_path_missing_file_includes_hint(tmp_path: Path):
    missing = tmp_path / "missing.docx"
    result = _validate_docx_path(str(missing))
    assert result is not None
    assert "does not exist" in result["error"]
    assert "hint" in result
    assert "create_document" in result["hint"]


def test_validate_docx_path_wrong_extension_includes_hint(tmp_path: Path):
    txt = tmp_path / "notes.txt"
    txt.write_text("hello", encoding="utf-8")
    result = _validate_docx_path(str(txt))
    assert result is not None
    assert "not a .docx" in result["error"]
    assert ".docx" in result["hint"]


def test_validate_docx_path_corrupt_file_includes_hint(tmp_path: Path):
    bad = tmp_path / "corrupt.docx"
    bad.write_bytes(b"not a zip")
    result = _validate_docx_path(str(bad))
    assert result is not None
    assert "not a valid Word document" in result["error"]
    assert "hint" in result


async def test_validate_docx_read_returns_hint_for_missing_file(tmp_path: Path):
    @validate_docx_read("filename")
    async def dummy(filename: str) -> dict[str, str]:
        return {"status": "ok"}

    result = await dummy(filename=str(tmp_path / "nope.docx"))
    assert "error" in result
    assert "hint" in result


async def test_validate_file_write_returns_hint_on_unwriteable_target(tmp_path: Path):
    from docx import Document

    doc_path = tmp_path / "locked.docx"
    Document().save(str(doc_path))

    @validate_file_write("filename")
    async def dummy(filename: str) -> dict[str, str]:
        return {"status": "ok"}

    import word_mcp.validation.document_validators as dv

    orig = dv._check_file_writeable
    dv._check_file_writeable = lambda _p: (False, "Permission denied")
    try:
        result = await dummy(filename=str(doc_path))
    finally:
        dv._check_file_writeable = orig

    assert "error" in result
    assert "hint" in result
