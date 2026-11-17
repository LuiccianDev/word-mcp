"""Tests for opt-in compound tool mode (WORD_MCP_TOOL_MODE=compound)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastmcp.exceptions import ToolError

from word_mcp.server import create_server


EXPECTED_COMPOUND_TOOLS = {
    "build_document",
    "create_from_markdown",
    "batch_add_content",
    "word_document",
    "word_content",
    "word_format",
    "word_protection",
    "word_footnote",
    "word_property",
}


async def _tool_names(server: Any) -> set[str]:
    return {tool.name for tool in await server.list_tools()}


async def _call_tool(server: Any, name: str, arguments: dict[str, Any]) -> Any:
    result = await server.call_tool(name, arguments)
    if hasattr(result, "structured_content") and result.structured_content is not None:
        return result.structured_content
    return result


@pytest.mark.asyncio
async def test_default_tool_mode_is_granular(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("WORD_MCP_TOOL_MODE", raising=False)
    server = create_server()
    names = await _tool_names(server)
    assert len(names) == 40
    assert "create_document" in names
    assert "word_document" not in names


@pytest.mark.asyncio
async def test_explicit_granular_mode(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "granular")
    server = create_server()
    names = await _tool_names(server)
    assert len(names) == 40


@pytest.mark.asyncio
async def test_compound_mode_registers_nine_tools(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    names = await _tool_names(server)
    assert names == EXPECTED_COMPOUND_TOOLS


def test_invalid_tool_mode_raises_value_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "invalid_mode")
    with pytest.raises(ValueError, match="WORD_MCP_TOOL_MODE"):
        create_server()


@pytest.mark.asyncio
async def test_compound_unknown_action_raises_tool_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    with pytest.raises(ToolError, match="Unknown action"):
        await _call_tool(server, "word_document", {"action": "nonexistent"})


@pytest.mark.asyncio
async def test_compound_unknown_param_raises_tool_error(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    server = create_server()
    with pytest.raises(ToolError, match="unknown parameter"):
        await _call_tool(
            server,
            "word_document",
            {
                "action": "create",
                "params": {
                    "filename": str(tmp_path / "typo.docx"),
                    "titel": "Typo Title",
                },
            },
        )


@pytest.mark.asyncio
async def test_compound_missing_required_param_raises_tool_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    with pytest.raises(ToolError, match="missing required parameter"):
        await _call_tool(server, "word_document", {"action": "create", "params": {}})


@pytest.mark.asyncio
async def test_compound_end_to_end_document_workflow(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    server = create_server()
    doc_path = str(tmp_path / "compound_test.docx")

    created = await _call_tool(
        server,
        "word_document",
        {"action": "create", "params": {"filename": doc_path, "title": "Compound"}},
    )
    assert created["status"] == "success"

    heading = await _call_tool(
        server,
        "word_content",
        {
            "action": "add_heading",
            "params": {"filename": doc_path, "text": "Overview", "level": 1},
        },
    )
    assert heading["status"] == "success"

    para = await _call_tool(
        server,
        "word_content",
        {
            "action": "add_paragraph",
            "params": {
                "filename": doc_path,
                "text": "Generated via compound tool mode.",
            },
        },
    )
    assert para["status"] == "success"

    props = await _call_tool(
        server,
        "word_property",
        {
            "action": "set_properties",
            "params": {"file_path": doc_path, "author": "Word MCP"},
        },
    )
    assert props["status"] == "success"

    text_res = await _call_tool(
        server,
        "word_document",
        {"action": "text", "params": {"filename": doc_path}},
    )
    assert "Generated via compound tool mode." in text_res["text"]


@pytest.mark.asyncio
async def test_compound_unknown_action_suggests_close_match(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    with pytest.raises(ToolError, match="Did you mean 'create'"):
        await _call_tool(server, "word_document", {"action": "creat", "params": {}})


@pytest.mark.asyncio
async def test_compound_unknown_param_suggests_close_match(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    with pytest.raises(ToolError, match="Did you mean 'filename'"):
        await _call_tool(
            server,
            "word_document",
            {"action": "info", "params": {"filname": "doc.docx"}},
        )


@pytest.mark.asyncio
async def test_compound_missing_required_param_lists_expected_signature(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    with pytest.raises(ToolError, match="Required parameters:"):
        await _call_tool(
            server,
            "word_content",
            {"action": "add_heading", "params": {"filename": "doc.docx"}},
        )


@pytest.mark.asyncio
async def test_compound_type_mismatch_raises_clear_tool_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    with pytest.raises(ToolError, match="Invalid type for 'level'"):
        await _call_tool(
            server,
            "word_content",
            {
                "action": "add_heading",
                "params": {"filename": "doc.docx", "text": "Hi", "level": "one"},
            },
        )


@pytest.mark.asyncio
async def test_compound_help_action_returns_action_schemas(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    server = create_server()
    overview = await _call_tool(server, "word_document", {"action": "help"})
    assert overview["tool"] == "word_document"
    assert "create" in overview["actions"]

    detail = await _call_tool(
        server,
        "word_document",
        {"action": "help", "params": {"action": "create"}},
    )
    assert detail["action"] == "create"
    assert "filename" in detail["required"]


@pytest.mark.asyncio
async def test_compound_coerces_valid_stringified_int_without_type_error(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    server = create_server()
    doc_path = str(tmp_path / "coerce_test.docx")
    await _call_tool(
        server, "word_document", {"action": "create", "params": {"filename": doc_path}}
    )
    res = await _call_tool(
        server,
        "word_content",
        {
            "action": "add_table",
            "params": {"filename": doc_path, "rows": "2", "cols": "2"},
        },
    )
    assert res["status"] == "success"


@pytest.mark.asyncio
async def test_compound_help_not_shadowed_by_domain_action() -> None:
    from word_mcp.tools.compound_tools import _dispatch

    async def fake_help() -> dict[str, str]:
        return {"shadowed": "yes"}

    result = await _dispatch("mock_group", {"help": fake_help}, "help", None)
    assert "actions" in result
    assert "shadowed" not in result


@pytest.mark.asyncio
async def test_compound_word_document_read_range(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("WORD_MCP_TOOL_MODE", "compound")
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    server = create_server()
    doc_path = str(tmp_path / "range_compound.docx")
    await _call_tool(
        server,
        "build_document",
        {
            "filename": doc_path,
            "elements": [
                {"type": "heading", "text": "Overview", "level": 1},
                {"type": "paragraph", "text": "Summary body."},
                {"type": "heading", "text": "Next", "level": 1},
                {"type": "paragraph", "text": "Next body."},
            ],
        },
    )
    help_res = await _call_tool(
        server,
        "word_document",
        {"action": "help", "params": {"action": "read_range"}},
    )
    assert help_res["action"] == "read_range"
    res = await _call_tool(
        server,
        "word_document",
        {
            "action": "read_range",
            "params": {"filename": doc_path, "section_heading": "Overview"},
        },
    )
    assert res["status"] == "success"
    assert [p["text"] for p in res["paragraphs"]] == ["Overview", "Summary body."]
