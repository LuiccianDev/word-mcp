"""Behavioral tests for sandbox enforcement, OperationError unification, atomic I/O, and compound dispatch."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import patch

import docx as real_docx
import pytest
from docx import Document
from fastmcp.exceptions import ToolError
from PIL import Image

from word_mcp.core.document_context import atomic_write_target, document_context
from word_mcp.server import create_server
from word_mcp.tools import (
    batch_tools,
    content_tools,
    document_tools,
    property_tools,
    protection_tools,
)


def _assert_canonical_error(
    payload: dict[str, Any],
    *,
    error_type: str,
    recoverable: bool,
) -> None:
    assert payload["status"] == "error"
    assert payload["error_type"] == error_type
    assert payload["recoverable"] is recoverable
    assert isinstance(payload["message"], str) and payload["message"]
    assert isinstance(payload["suggestion"], str) and payload["suggestion"]


async def _call_tool(server: Any, name: str, arguments: dict[str, Any]) -> Any:
    result = await server.call_tool(name, arguments)
    if hasattr(result, "structured_content") and result.structured_content is not None:
        return result.structured_content
    return result


@pytest.mark.asyncio
async def test_sandbox_enforced_across_primary_and_secondary_paths(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    allowed_dir = tmp_path / "allowed"
    outside_dir = tmp_path / "outside"
    allowed_dir.mkdir()
    outside_dir.mkdir()
    monkeypatch.setenv("WORD_MCP_DIRS", str(allowed_dir))

    inside_doc = allowed_dir / "inside.docx"
    Document().save(inside_doc)

    outside_doc = outside_dir / "secret.docx"
    Document().save(outside_doc)

    outside_img = outside_dir / "escape.png"
    Image.new("RGB", (10, 10), color="blue").save(outside_img)

    read_res = await document_tools.get_document_text(str(outside_doc))
    _assert_canonical_error(
        read_res, error_type="security_violation", recoverable=False
    )
    assert "WORD_MCP_DIRS" in read_res["suggestion"]

    traversal_res = await document_tools.get_document_text(
        str(allowed_dir / ".." / "outside" / "secret.docx")
    )
    _assert_canonical_error(
        traversal_res, error_type="security_violation", recoverable=False
    )

    copy_res = await document_tools.copy_document(
        str(inside_doc), str(outside_dir / "copy.docx")
    )
    _assert_canonical_error(
        copy_res, error_type="security_violation", recoverable=False
    )

    merge_res = await document_tools.merge_documents(
        str(allowed_dir / "merged.docx"),
        [str(inside_doc), str(outside_doc)],
    )
    _assert_canonical_error(
        merge_res, error_type="security_violation", recoverable=False
    )

    pdf_res = await document_tools.convert_to_pdf(
        str(inside_doc), str(outside_dir / "out.pdf")
    )
    _assert_canonical_error(pdf_res, error_type="security_violation", recoverable=False)

    pic_res = await content_tools.add_picture(str(inside_doc), str(outside_img))
    _assert_canonical_error(pic_res, error_type="security_violation", recoverable=False)

    list_res = await document_tools.list_available_documents(directory=str(outside_dir))
    _assert_canonical_error(
        list_res, error_type="security_violation", recoverable=False
    )

    batch_img_res = await batch_tools.build_document(
        filename=str(allowed_dir / "batch.docx"),
        elements=[{"type": "image", "path": str(outside_img)}],
    )
    _assert_canonical_error(
        batch_img_res, error_type="security_violation", recoverable=False
    )


@pytest.mark.asyncio
async def test_sandbox_rejects_empty_whitespace_and_nul_paths(
    tmp_path: Path,
) -> None:
    for bad_path in ("", "   ", f"{tmp_path / 'bad'}\x00.docx"):
        res = await document_tools.get_document_text(bad_path)
        _assert_canonical_error(res, error_type="validation", recoverable=False)


@pytest.mark.asyncio
async def test_list_available_documents_scopes_to_requested_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dir_a = tmp_path / "a"
    dir_b = tmp_path / "b"
    dir_a.mkdir()
    dir_b.mkdir()
    Document().save(dir_a / "alpha.docx")
    Document().save(dir_b / "beta.docx")
    monkeypatch.setenv("WORD_MCP_DIRS", f"{dir_a},{dir_b}")

    scoped = await document_tools.list_available_documents(directory=str(dir_b))
    assert scoped["status"] == "success"
    assert [doc["name"] for doc in scoped["documents"]] == ["beta.docx"]


@pytest.mark.asyncio
async def test_unified_operation_error_shape_across_tools(tmp_path: Path) -> None:
    doc_path = tmp_path / "sample.docx"
    doc = Document()
    doc.add_paragraph("Hello world.")
    doc.save(doc_path)

    empty_props = await property_tools.set_core_properties(str(doc_path))
    _assert_canonical_error(empty_props, error_type="validation", recoverable=False)

    empty_sections = await protection_tools.add_restricted_editing(
        str(doc_path), "pw", []
    )
    _assert_canonical_error(empty_sections, error_type="validation", recoverable=False)

    not_found_replace = await content_tools.search_and_replace(
        str(doc_path), "absent_token", "replacement"
    )
    _assert_canonical_error(
        not_found_replace, error_type="validation", recoverable=False
    )

    missing_batch = await batch_tools.batch_add_content(
        elements=[{"type": "paragraph", "text": "x"}],
        filename=str(tmp_path / "missing.docx"),
    )
    _assert_canonical_error(
        missing_batch, error_type="file_not_found", recoverable=False
    )

    conflict_batch = await batch_tools.build_document(
        filename=str(tmp_path / "a.docx"),
        file_path=str(tmp_path / "b.docx"),
    )
    _assert_canonical_error(conflict_batch, error_type="validation", recoverable=False)

    rel_alias = await batch_tools.build_document(
        file_path=str(tmp_path / "subdir" / "alias_no_ext"),
        elements=[{"type": "paragraph", "text": "Normalized"}],
    )
    assert rel_alias["status"] == "success"
    assert rel_alias["filename"].endswith("alias_no_ext.docx")
    assert Path(rel_alias["filename"]).is_file()


@pytest.mark.asyncio
async def test_read_tools_open_document_only_once(tmp_path: Path) -> None:
    doc_path = tmp_path / "single_open.docx"
    doc = Document()
    doc.add_heading("Section One", level=1)
    doc.add_paragraph("Body line.")
    doc.save(doc_path)

    open_count = 0
    real_factory = real_docx.Document

    def counting_document(*args: Any, **kwargs: Any) -> Any:
        nonlocal open_count
        open_count += 1
        return real_factory(*args, **kwargs)

    with patch(
        "word_mcp.core.document_context.Document", side_effect=counting_document
    ):
        text_res = await document_tools.get_document_text(str(doc_path))
        assert text_res["status"] == "success"
        assert open_count == 1

        open_count = 0
        outline_res = await document_tools.get_document_outline(str(doc_path))
        assert len(outline_res["paragraphs"]) == 2
        assert open_count == 1

        open_count = 0
        range_res = await document_tools.read_document_range(str(doc_path))
        assert range_res["status"] == "success"
        assert open_count == 1


def test_atomic_write_and_document_context_rollback_on_exception(
    tmp_path: Path,
) -> None:
    doc_path = tmp_path / "atomic.docx"
    doc = Document()
    doc.add_paragraph("Original content")
    doc.save(doc_path)

    with pytest.raises(RuntimeError, match="abort"):
        with document_context(str(doc_path), mode="write") as opened:
            opened.add_paragraph("Unsaved change")
            raise RuntimeError("abort")

    reloaded = Document(doc_path)
    assert [p.text for p in reloaded.paragraphs] == ["Original content"]
    leftovers = [p for p in tmp_path.iterdir() if p.name != "atomic.docx"]
    assert leftovers == []

    with pytest.raises(ValueError, match="mid-write"):
        with atomic_write_target(doc_path) as tmp_target:
            tmp_target.write_bytes(b"partial-corrupt")
            raise ValueError("mid-write")

    reloaded_after = Document(doc_path)
    assert [p.text for p in reloaded_after.paragraphs] == ["Original content"]
    leftovers_after = [p for p in tmp_path.iterdir() if p.name != "atomic.docx"]
    assert leftovers_after == []


@pytest.mark.asyncio
async def test_verify_document_on_unprotected_and_encrypted_docx(
    tmp_path: Path,
) -> None:
    doc_path = tmp_path / "protected.docx"
    doc = Document()
    doc.add_paragraph("Confidential body")
    doc.save(doc_path)

    plain_verify = await protection_tools.verify_document(str(doc_path))
    assert plain_verify["status"] == "success"
    assert plain_verify["success"] is False

    enc_res = await protection_tools.protect_document(str(doc_path), "s3cr3t")
    assert enc_res["status"] == "success"
    assert enc_res["success"] is True

    enc_verify_no_pw = await protection_tools.verify_document(str(doc_path))
    assert enc_verify_no_pw["status"] == "success"
    assert enc_verify_no_pw["success"] is True

    enc_verify_bad_pw = await protection_tools.verify_document(
        str(doc_path), password="wrong"
    )
    assert enc_verify_bad_pw["status"] == "success"
    assert enc_verify_bad_pw["success"] is False

    enc_verify_ok = await protection_tools.verify_document(
        str(doc_path), password="s3cr3t"
    )
    assert enc_verify_ok["status"] == "success"
    assert enc_verify_ok["success"] is True

    bad_unprotect = await protection_tools.unprotect_document(str(doc_path), "wrong")
    _assert_canonical_error(bad_unprotect, error_type="validation", recoverable=False)

    ok_unprotect = await protection_tools.unprotect_document(str(doc_path), "s3cr3t")
    assert ok_unprotect["status"] == "success"
    assert ok_unprotect["success"] is True


@pytest.mark.asyncio
async def test_compound_rich_help_required_signatures_aliases_and_bool_guard(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOLS", "compound")
    server = create_server()

    doc_help = await _call_tool(server, "word_document", {"action": "help"})
    assert doc_help["aliases"] == {"file_path": "filename"}
    assert "required" in doc_help["actions"]["create"]
    assert "optional" in doc_help["actions"]["create"]

    expected_required_filename = [
        ("word_content", "add_hyperlink"),
        ("word_content", "add_bookmark"),
        ("word_format", "add_header"),
        ("word_format", "add_footer"),
        ("word_property", "get_properties"),
        ("word_property", "set_properties"),
        ("word_property", "set_page_layout"),
    ]
    for group, action in expected_required_filename:
        group_help = await _call_tool(server, group, {"action": "help"})
        assert "filename" in group_help["actions"][action]["required"]

    doc_path = tmp_path / "compound_alias.docx"
    Document().save(doc_path)

    header_res = await _call_tool(
        server,
        "word_format",
        {
            "action": "add_header",
            "params": {"file_path": str(doc_path), "text": "Alias Header"},
        },
    )
    assert header_res["status"] == "success"

    with pytest.raises(ToolError, match="not both"):
        await _call_tool(
            server,
            "word_format",
            {
                "action": "add_header",
                "params": {
                    "filename": str(doc_path),
                    "file_path": str(doc_path),
                    "text": "Conflict",
                },
            },
        )

    with pytest.raises(ToolError, match="Invalid type for 'level'"):
        await _call_tool(
            server,
            "word_content",
            {
                "action": "add_heading",
                "params": {"filename": str(doc_path), "text": "Title", "level": True},
            },
        )

    with pytest.raises(ToolError, match="Invalid type for 'width'"):
        await _call_tool(
            server,
            "word_content",
            {
                "action": "add_picture",
                "params": {
                    "filename": str(doc_path),
                    "image_path": str(tmp_path / "x.png"),
                    "width": True,
                },
            },
        )


@pytest.mark.asyncio
async def test_thermos_edge_cases_batch_image_positional_dir_and_corrupt_verify(
    tmp_path: Path,
) -> None:
    from word_mcp.tools.compound_tools import _GROUP_SPECS, _get_action_spec

    for spec in _GROUP_SPECS:
        action_spec = _get_action_spec(
            spec.example_action, spec.actions[spec.example_action]
        )
        bound = action_spec.bind_and_validate(spec.name, spec.example_params)
        assert bound

    doc_path = tmp_path / "valid.docx"
    Document().save(doc_path)

    bad_img_nul = await batch_tools.build_document(
        filename=str(tmp_path / "out.docx"),
        elements=[{"type": "image", "path": f"{tmp_path / 'bad'}\x00.png"}],
    )
    _assert_canonical_error(bad_img_nul, error_type="validation", recoverable=False)

    missing_img = await batch_tools.build_document(
        filename=str(tmp_path / "out2.docx"),
        elements=[{"type": "image", "path": str(tmp_path / "missing.png")}],
    )
    _assert_canonical_error(missing_img, error_type="file_not_found", recoverable=False)

    pos_conflict = await batch_tools.build_document(
        str(tmp_path / "a.docx"),
        file_path=str(tmp_path / "b.docx"),
    )
    _assert_canonical_error(pos_conflict, error_type="validation", recoverable=False)

    dir_as_docx = tmp_path / "folder.docx"
    dir_as_docx.mkdir()
    dir_read = await document_tools.get_document_text(str(dir_as_docx))
    _assert_canonical_error(dir_read, error_type="validation", recoverable=False)

    empty_docx = tmp_path / "empty.docx"
    empty_docx.write_bytes(b"")
    verify_empty = await protection_tools.verify_document(str(empty_docx))
    _assert_canonical_error(
        verify_empty, error_type="document_processing", recoverable=True
    )


@pytest.mark.asyncio
async def test_security_unc_device_and_reserved_names_rejected_before_resolve(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))

    for bad_path in (
        r"\\attacker.example\share\evil.docx",
        "//attacker.example/share/evil.docx",
        r"\\.\pipe\word_mcp_probe",
        r"\\?\C:\temp\evil.docx",
    ):
        res = await document_tools.get_document_text(bad_path)
        _assert_canonical_error(res, error_type="security_violation", recoverable=False)

    for reserved_name in ("CON.docx", "NUL.docx", "COM1.docx"):
        res = await document_tools.create_document(str(tmp_path / reserved_name))
        _assert_canonical_error(res, error_type="validation", recoverable=False)


@pytest.mark.asyncio
async def test_ooxml_zip_bomb_rejected_before_document_open(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import zipfile

    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))
    bomb_path = tmp_path / "bomb.docx"
    with zipfile.ZipFile(bomb_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", "A" * (2 * 1024 * 1024))

    res = await document_tools.get_document_text(str(bomb_path))
    _assert_canonical_error(res, error_type="document_processing", recoverable=True)


@pytest.mark.asyncio
async def test_protection_empty_password_signature_tamper_and_sidecar_hardening(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from word_mcp.core.protection import remove_protection_info, verify_password

    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))

    enc_doc = tmp_path / "enc.docx"
    Document().save(enc_doc)
    await protection_tools.protect_document(str(enc_doc), "TopSecret1!")

    empty_enc = await protection_tools.verify_document(str(enc_doc), password="")
    assert empty_enc["status"] == "success"
    assert empty_enc["success"] is False

    restricted_doc = tmp_path / "restricted.docx"
    Document().save(restricted_doc)
    await protection_tools.add_restricted_editing(
        str(restricted_doc), "Restr1!", ["Section 1"]
    )
    empty_restr = await protection_tools.verify_document(
        str(restricted_doc), password=""
    )
    assert empty_restr["status"] == "success"
    assert empty_restr["success"] is False

    ok_remove, msg_remove = remove_protection_info(str(restricted_doc), password="")
    assert ok_remove is False
    assert msg_remove == "Incorrect password"
    ok_none, _ = remove_protection_info(str(restricted_doc), password=None)
    assert ok_none is False

    assert verify_password("pw", "pbkdf2_sha256$999999999$salt$abcd") is False
    assert verify_password("pw", "pbkdf2_sha256$1$salt$abcd") is False

    signed_doc = tmp_path / "signed.docx"
    d = Document()
    d.add_paragraph("Contract Clause 1")
    tbl = d.add_table(rows=1, cols=2)
    tbl.cell(0, 0).text = "Amount"
    tbl.cell(0, 1).text = "$100"
    d.save(signed_doc)

    sign_res = await protection_tools.add_digital_signature(
        str(signed_doc), "Alice", reason="Approval"
    )
    assert sign_res["status"] == "success"

    verify_ok = await protection_tools.verify_document(str(signed_doc))
    assert verify_ok["status"] == "success"
    assert verify_ok["success"] is True

    tampered = Document(signed_doc)
    tampered.tables[0].cell(0, 1).text = "$999"
    tampered.save(signed_doc)

    verify_tampered = await protection_tools.verify_document(str(signed_doc))
    assert verify_tampered["status"] == "success"
    assert verify_tampered["success"] is False


@pytest.mark.asyncio
async def test_protection_sidecar_symlink_rejected(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import os

    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))
    doc_path = tmp_path / "linked.docx"
    Document().save(doc_path)
    outside_secret = tmp_path.parent / "outside_secret.json"
    outside_secret.write_text('{"type": "password"}', encoding="utf-8")
    meta_link = tmp_path / "linked.protection"

    try:
        os.symlink(outside_secret, meta_link)
    except (OSError, NotImplementedError):
        pytest.skip("Symlink creation requires elevated privileges on this host")

    res = await protection_tools.verify_document(str(doc_path))
    _assert_canonical_error(res, error_type="security_violation", recoverable=False)


@pytest.mark.asyncio
async def test_hyperlink_scheme_allowlist_and_unique_bookmark_ids(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docx.oxml.ns import qn

    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))
    doc_path = tmp_path / "links.docx"
    doc = Document()
    doc.add_paragraph("Visit link")
    doc.save(doc_path)

    for unsafe in (
        "javascript:alert(1)",
        "file:///etc/passwd",
        r"\\attacker\share\doc",
        "smb://attacker/share",
    ):
        bad = await content_tools.add_hyperlink(str(doc_path), 0, "Bad", url=unsafe)
        _assert_canonical_error(bad, error_type="validation", recoverable=False)

    ok_https = await content_tools.add_hyperlink(
        str(doc_path), 0, "Safe", url="https://example.com"
    )
    assert ok_https["status"] == "success"
    ok_mailto = await content_tools.add_hyperlink(
        str(doc_path), 0, "Mail", url="mailto:security@example.com"
    )
    assert ok_mailto["status"] == "success"

    await content_tools.add_bookmark(str(doc_path), 0, "BM1")
    await content_tools.add_bookmark(str(doc_path), 0, "BM2")
    reloaded = Document(doc_path)
    ids = [el.get(qn("w:id")) for el in reloaded.element.xpath(".//w:bookmarkStart")]
    assert ids == ["0", "1"]


@pytest.mark.asyncio
async def test_resource_caps_and_list_available_documents_hygiene(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from word_mcp.tools import footnote_tools

    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))
    doc_path = tmp_path / "bounds.docx"
    Document().save(doc_path)

    bad_table = await content_tools.add_table(str(doc_path), rows=50_000, cols=10)
    _assert_canonical_error(bad_table, error_type="validation", recoverable=False)

    bad_fn = await footnote_tools.customize_footnote_style(
        str(doc_path), start_number=100_000
    )
    _assert_canonical_error(bad_fn, error_type="validation", recoverable=False)

    (tmp_path / "folder.docx").mkdir()
    listed = await document_tools.list_available_documents(directory=str(tmp_path))
    assert listed["status"] == "success"
    assert [d["name"] for d in listed["documents"]] == ["bounds.docx"]


@pytest.mark.asyncio
async def test_core_format_search_replace_and_merge_preserve_structure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("WORD_MCP_DIRS", str(tmp_path))

    fmt_path = tmp_path / "fmt.docx"
    d = Document()
    p = d.add_paragraph()
    r1 = p.add_run("Hello ")
    r1.bold = True
    p.add_run("world today")
    d.save(fmt_path)

    await content_tools.batch_edit_document(
        str(fmt_path),
        operations=[{"op": "replace_text", "old_text": "Hello", "new_text": "Hi"}],
    )
    from word_mcp.tools import format_tools

    await format_tools.format_text(str(fmt_path), 0, 3, 8, italic=True)
    reloaded_fmt = Document(fmt_path)
    assert reloaded_fmt.paragraphs[0].text == "Hi world today"
    assert reloaded_fmt.paragraphs[0].runs[0].bold is True
    assert reloaded_fmt.paragraphs[0].runs[1].italic is True

    split_path = tmp_path / "split.docx"
    sd = Document()
    sp = sd.add_paragraph()
    sp.add_run("{{CLIENT_")
    sp.add_run("NAME}} and {{CLIENT_")
    sp.add_run("NAME}}")
    sd.sections[0].header.paragraphs[0].text = "Header {{CLIENT_NAME}}"
    sd.add_section()
    sd.save(split_path)

    rep_res = await content_tools.search_and_replace(
        str(split_path), "{{CLIENT_NAME}}", "Acme"
    )
    assert rep_res["status"] == "success"
    assert rep_res["details"]["replacements_made"] == 3
    reloaded_split = Document(split_path)
    assert reloaded_split.paragraphs[0].text == "Acme and Acme"
    assert reloaded_split.sections[0].header.paragraphs[0].text == "Header Acme"

    rep_expand = await content_tools.search_and_replace(
        str(split_path), "Acme", "Acme Corp"
    )
    assert rep_expand["status"] == "success"
    assert rep_expand["details"]["replacements_made"] == 3
    assert (
        Document(split_path).sections[0].header.paragraphs[0].text == "Header Acme Corp"
    )

    src1 = tmp_path / "src1.docx"
    s1 = Document()
    mp = s1.add_paragraph()
    mr1 = mp.add_run("Bold ")
    mr1.bold = True
    mr2 = mp.add_run("Italic")
    mr2.italic = True
    t = s1.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "Middle Table"
    s1.add_paragraph("After Table")
    s1.save(src1)

    merged_path = tmp_path / "merged.docx"
    merge_res = await document_tools.merge_documents(
        str(merged_path), [str(src1)], add_page_breaks=False
    )
    assert merge_res["status"] == "success"
    merged = Document(merged_path)
    body_tags = [child.tag.rsplit("}", 1)[-1] for child in merged.element.body]
    p_indices = [i for i, tag in enumerate(body_tags) if tag == "p"]
    tbl_indices = [i for i, tag in enumerate(body_tags) if tag == "tbl"]
    assert len(tbl_indices) == 1
    assert p_indices[0] < tbl_indices[0] < p_indices[-1]
    assert merged.paragraphs[0].runs[0].bold is True
    assert merged.paragraphs[0].runs[1].italic is True


def test_cli_sse_rejects_non_loopback_without_allow_remote(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import sys

    from word_mcp.__main__ import main

    monkeypatch.setattr(
        sys, "argv", ["luiccian-word-mcp", "--transport", "sse", "--host", "0.0.0.0"]
    )
    with pytest.raises(SystemExit, match="allow-remote"):
        main()
