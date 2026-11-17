"""Compact (compound) tool registrations for Word MCP.

When ``WORD_MCP_TOOLS=compound`` (the default, also aliased as ``compact``),
these 9 tools replace the 40 full/granular tools: the 3 high-efficiency
single-pass tools (``build_document``, ``create_from_markdown``,
``batch_add_content``) keep their full typed JSON schemas, while the remaining
37 atomic tools are consolidated into 6 action-dispatched domain tools
(``word_document``, ``word_content``, ``word_format``, ``word_protection``,
``word_footnote``, ``word_property``).
"""

from __future__ import annotations

import difflib
import inspect
from collections.abc import Awaitable, Callable
from typing import Any, get_type_hints

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import TypeAdapter, ValidationError as PydanticValidationError

from word_mcp.tools import (
    batch_tools,
    content_tools,
    document_tools,
    footnote_tools,
    format_tools,
    property_tools,
    protection_tools,
)


_PARAMS_NOTE = (
    "\nAll action parameters go inside `params`, e.g. "
    '{"action": "create", "params": {"filename": "/path/doc.docx"}}. '
    "Omit `params` for actions that take none, or use `action='help'` to inspect parameters."
)

_ToolFn = Callable[..., Awaitable[dict[str, Any]]]
_TYPE_ADAPTER_CACHE: dict[Any, TypeAdapter[Any]] = {}


def _format_type_name(annotation: Any) -> str:
    """Render a concise human-readable type name for error messages and help."""
    if annotation is inspect.Parameter.empty:
        return "Any"
    if isinstance(annotation, type):
        return annotation.__name__
    return str(annotation).replace("typing.", "")


def _get_resolved_hints(handler: _ToolFn) -> dict[str, Any]:
    """Resolve type hints safely, falling back to raw annotations on failure."""
    try:
        return get_type_hints(handler)
    except Exception:
        return {}


def _describe_action(handler: _ToolFn) -> dict[str, Any]:
    """Extract required/optional parameter metadata and summary for an action."""
    sig = inspect.signature(handler)
    hints = _get_resolved_hints(handler)
    required: dict[str, str] = {}
    optional: dict[str, dict[str, Any]] = {}

    for param in sig.parameters.values():
        ann = hints.get(param.name, param.annotation)
        type_str = _format_type_name(ann)
        if param.default is inspect.Parameter.empty:
            required[param.name] = type_str
        else:
            optional[param.name] = {"type": type_str, "default": param.default}

    doc = inspect.getdoc(handler) or ""
    summary = doc.splitlines()[0].strip() if doc else ""
    return {
        "description": summary,
        "required": required,
        "optional": optional,
    }


def _format_suggestion(target: str, candidates: list[str]) -> str:
    """Return a close-match hint string or empty string if none found."""
    matches = difflib.get_close_matches(target, candidates, n=2, cutoff=0.5)
    if not matches:
        return ""
    if len(matches) == 1:
        return f" Did you mean '{matches[0]}'?"
    return f" Did you mean one of: {', '.join(repr(m) for m in matches)}?"


def _validate_param_types(
    group: str, action: str, handler: _ToolFn, kwargs: dict[str, Any]
) -> None:
    """Validate parameter values against resolved annotations and raise ToolError."""
    sig = inspect.signature(handler)
    hints = _get_resolved_hints(handler)
    for name, value in kwargs.items():
        param = sig.parameters.get(name)
        if param is None:
            continue
        ann = hints.get(name, param.annotation)
        if ann is inspect.Parameter.empty or ann is Any:
            continue
        adapter = _TYPE_ADAPTER_CACHE.get(ann)
        if adapter is None:
            try:
                adapter = TypeAdapter(ann)
            except Exception:
                continue
            _TYPE_ADAPTER_CACHE[ann] = adapter
        try:
            kwargs[name] = adapter.validate_python(value)
        except PydanticValidationError as exc:
            expected = _format_type_name(ann)
            raise ToolError(
                f"{group} action '{action}': Invalid type for '{name}': "
                f"expected {expected}, got {type(value).__name__}"
            ) from exc


class _Params(dict[str, Any]):
    """Action params wrapper that catches missing required and unknown keys."""

    def __init__(
        self,
        group: str,
        action: str,
        params: dict[str, Any],
        schema: dict[str, Any],
    ) -> None:
        super().__init__(params)
        self._where = f"{group} action '{action}'"
        self._schema = schema
        self._read: set[str] = set()

    def __getitem__(self, key: str) -> Any:
        self._read.add(key)
        return super().__getitem__(key)

    def __missing__(self, key: str) -> Any:
        req = ", ".join(
            f"{name} ({tp})" for name, tp in self._schema["required"].items()
        ) or "none"
        opt = ", ".join(self._schema["optional"]) or "none"
        raise ToolError(
            f"{self._where}: missing required parameter '{key}'. "
            f"Required parameters: {req}. Optional parameters: {opt}."
        )

    def __contains__(self, key: object) -> bool:
        if isinstance(key, str):
            self._read.add(key)
        return super().__contains__(key)

    def check_all_read(self) -> None:
        """Raise ToolError if any caller-supplied key was not consumed."""
        unused = sorted(set(self) - self._read)
        if unused:
            valid = list(self._schema["required"]) + list(self._schema["optional"])
            hints = "".join(_format_suggestion(key, valid) for key in unused)
            valid_str = ", ".join(valid) or "none"
            raise ToolError(
                f"{self._where}: unknown parameter(s) {unused}.{hints} "
                f"Valid parameters: {valid_str}."
            )


def _handle_help(
    group: str, actions: dict[str, _ToolFn], params: dict[str, Any] | None
) -> dict[str, Any]:
    """Return overview or per-action parameter schema for a compound tool."""
    if params is not None and not isinstance(params, dict):
        raise ToolError(
            f"{group} action 'help': `params` must be an object/dict, "
            f"got {type(params).__name__}"
        )
    if params and "action" in params:
        target = params["action"]
        if not isinstance(target, str) or target not in actions:
            suggestion = (
                _format_suggestion(target, sorted(actions))
                if isinstance(target, str)
                else ""
            )
            valid = ", ".join(sorted(actions))
            raise ToolError(
                f"Unknown action '{target}' for {group}.{suggestion} Valid: {valid}"
            )
        return {"tool": group, "action": target, **_describe_action(actions[target])}

    return {
        "tool": group,
        "actions": {
            name: _describe_action(fn)["description"]
            for name, fn in sorted(actions.items())
        },
        "hint": "Call with action='help' and params={'action': '<name>'} for parameter details.",
    }


async def _dispatch(
    group: str,
    actions: dict[str, _ToolFn],
    action: str,
    params: dict[str, Any] | None,
) -> dict[str, Any]:
    """Validate action/params against the underlying tool signature and invoke it."""
    if action == "help":
        return _handle_help(group, actions, params)

    handler = actions.get(action)
    if handler is None:
        valid = ", ".join(sorted(actions))
        suggestion = _format_suggestion(action, sorted(actions))
        raise ToolError(
            f"Unknown action '{action}' for {group}.{suggestion} "
            f"Valid: {valid}. Call with action='help' for parameter details."
        )

    if params is not None and not isinstance(params, dict):
        raise ToolError(
            f"{group} action '{action}': `params` must be an object/dict, "
            f"got {type(params).__name__}"
        )

    schema = _describe_action(handler)
    validated = _Params(group, action, params or {}, schema)
    sig = inspect.signature(handler)
    unknown = sorted(set(validated) - set(sig.parameters))
    if unknown:
        valid_params: list[str] = list(schema["required"]) + list(schema["optional"])
        hints = "".join(_format_suggestion(key, valid_params) for key in unknown)
        valid_str = ", ".join(valid_params) or "none"
        raise ToolError(
            f"{group} action '{action}': unknown parameter(s) {unknown}.{hints} "
            f"Valid parameters: {valid_str}."
        )

    kwargs: dict[str, Any] = {}
    for param in sig.parameters.values():
        if param.default is inspect.Parameter.empty:
            kwargs[param.name] = validated[param.name]
        elif param.name in validated:
            kwargs[param.name] = validated[param.name]

    validated.check_all_read()
    _validate_param_types(group, action, handler, kwargs)
    return await handler(**kwargs)


_DOCUMENT_ACTIONS: dict[str, _ToolFn] = {
    "create": document_tools.create_document,
    "copy": document_tools.copy_document,
    "info": document_tools.get_document_info,
    "text": document_tools.get_document_text,
    "outline": document_tools.get_document_outline,
    "list": document_tools.list_available_documents,
    "merge": document_tools.merge_documents,
    "paragraph_text": document_tools.get_paragraph_text_from_document,
    "find_text": document_tools.find_text_in_document,
    "to_pdf": document_tools.convert_to_pdf,
    "read_range": document_tools.read_document_range,
    "get_tables": document_tools.get_document_tables,
}

_CONTENT_ACTIONS: dict[str, _ToolFn] = {
    "add_heading": content_tools.add_heading,
    "add_paragraph": content_tools.add_paragraph,
    "add_picture": content_tools.add_picture,
    "add_table": content_tools.add_table,
    "add_page_break": content_tools.add_page_break,
    "delete_paragraph": content_tools.delete_paragraph,
    "search_and_replace": content_tools.search_and_replace,
    "add_toc": content_tools.add_table_of_contents,
    "add_hyperlink": content_tools.add_hyperlink,
    "add_bookmark": content_tools.add_bookmark,
    "batch_edit": content_tools.batch_edit_document,
}

_FORMAT_ACTIONS: dict[str, _ToolFn] = {
    "format_text": format_tools.format_text,
    "create_style": format_tools.create_custom_style,
    "format_table": format_tools.format_table,
    "add_header": format_tools.add_header,
    "add_footer": format_tools.add_footer,
}

_PROTECTION_ACTIONS: dict[str, _ToolFn] = {
    "protect": protection_tools.protect_document,
    "unprotect": protection_tools.unprotect_document,
    "restrict_editing": protection_tools.add_restricted_editing,
    "add_signature": protection_tools.add_digital_signature,
    "verify": protection_tools.verify_document,
}

_FOOTNOTE_ACTIONS: dict[str, _ToolFn] = {
    "add_footnote": footnote_tools.add_footnote_to_document,
    "add_endnote": footnote_tools.add_endnote_to_document,
    "convert_to_endnotes": footnote_tools.convert_footnotes_to_endnotes_in_document,
    "customize_style": footnote_tools.customize_footnote_style,
}

_PROPERTY_ACTIONS: dict[str, _ToolFn] = {
    "get_properties": property_tools.get_core_properties,
    "set_properties": property_tools.set_core_properties,
    "set_page_layout": property_tools.set_page_layout,
}


def register_compound_tools(mcp: FastMCP) -> None:
    """Register the 9-tool hybrid compound surface on the FastMCP server."""
    write = ToolAnnotations(readOnlyHint=False)
    destructive = ToolAnnotations(readOnlyHint=False, destructiveHint=True)

    # Keep the 3 high-efficiency single-pass tools with full typed schemas
    mcp.tool(annotations=write)(batch_tools.build_document)
    mcp.tool(annotations=write)(batch_tools.create_from_markdown)
    mcp.tool(annotations=write)(batch_tools.batch_add_content)

    @mcp.tool(
        annotations=write,
        description=(
            "Document lifecycle, inspection, search, merge, and PDF export. "
            f"Actions: {', '.join(_DOCUMENT_ACTIONS)}." + _PARAMS_NOTE
        ),
    )
    async def word_document(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch("word_document", _DOCUMENT_ACTIONS, action, params)

    @mcp.tool(
        annotations=destructive,
        description=(
            "Add, update, search/replace, or delete document content elements. "
            f"Actions: {', '.join(_CONTENT_ACTIONS)}." + _PARAMS_NOTE
        ),
    )
    async def word_content(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch("word_content", _CONTENT_ACTIONS, action, params)

    @mcp.tool(
        annotations=write,
        description=(
            "Format text, tables, custom styles, headers, and footers. "
            f"Actions: {', '.join(_FORMAT_ACTIONS)}." + _PARAMS_NOTE
        ),
    )
    async def word_format(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch("word_format", _FORMAT_ACTIONS, action, params)

    @mcp.tool(
        annotations=write,
        description=(
            "Encrypt, unprotect, restrict editing, add signature marker, or verify. "
            f"Actions: {', '.join(_PROTECTION_ACTIONS)}." + _PARAMS_NOTE
        ),
    )
    async def word_protection(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch("word_protection", _PROTECTION_ACTIONS, action, params)

    @mcp.tool(
        annotations=write,
        description=(
            "Add footnotes/endnotes, convert footnotes, or customize note style. "
            f"Actions: {', '.join(_FOOTNOTE_ACTIONS)}." + _PARAMS_NOTE
        ),
    )
    async def word_footnote(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch("word_footnote", _FOOTNOTE_ACTIONS, action, params)

    @mcp.tool(
        annotations=write,
        description=(
            "Read/update core document properties and configure page layout. "
            f"Actions: {', '.join(_PROPERTY_ACTIONS)}." + _PARAMS_NOTE
        ),
    )
    async def word_property(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch("word_property", _PROPERTY_ACTIONS, action, params)
