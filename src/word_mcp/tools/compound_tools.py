"""Compact (compound) tool registrations for Word MCP.

When ``WORD_MCP_TOOLS=compound`` (the default),
these 9 tools replace the 40 full/granular tools: the 3 high-efficiency
single-pass tools (``build_document``, ``create_from_markdown``,
``batch_add_content``) keep their full typed JSON schemas, while the remaining
37 granular tools plus 3 compound-only actions (40 actions total) are
consolidated into 6 action-dispatched domain tools
(``word_document``, ``word_content``, ``word_format``, ``word_protection``,
``word_footnote``, ``word_property``).
"""

from __future__ import annotations

import difflib
import inspect
import types
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from typing import Any, Union, get_args, get_origin, get_type_hints

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


_ToolFn = Callable[..., Awaitable[dict[str, Any]]]
_PARAM_ALIASES: dict[str, str] = {"file_path": "filename"}
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


def _format_suggestion(target: str, candidates: list[str]) -> str:
    """Return a close-match hint string or empty string if none found."""
    matches = difflib.get_close_matches(target, candidates, n=2, cutoff=0.5)
    if not matches:
        return ""
    if len(matches) == 1:
        return f" Did you mean '{matches[0]}'?"
    return f" Did you mean one of: {', '.join(repr(m) for m in matches)}?"


def _expects_numeric_without_bool(annotation: Any) -> bool:
    """Return True when annotation expects int/float without allowing bool."""
    if annotation in (int, float):
        return True
    origin = get_origin(annotation)
    if origin in (Union, types.UnionType):
        non_none = [arg for arg in get_args(annotation) if arg is not type(None)]
        return any(arg in (int, float) for arg in non_none) and bool not in non_none
    return False


def _has_disallowed_bool(annotation: Any, value: Any) -> bool:
    """Reject bool passed into int/float fields (and lists of int/float)."""
    if _expects_numeric_without_bool(annotation):
        return isinstance(value, bool)
    origin = get_origin(annotation)
    if origin in (Union, types.UnionType):
        return any(
            _has_disallowed_bool(arg, value)
            for arg in get_args(annotation)
            if arg is not type(None)
        )
    if origin in (list, tuple) and isinstance(value, (list, tuple)):
        elem_args = get_args(annotation)
        if elem_args and _expects_numeric_without_bool(elem_args[0]):
            return any(isinstance(item, bool) for item in value)
    return False


def _get_type_adapter(annotation: Any) -> TypeAdapter[Any] | None:
    """Return a cached Pydantic TypeAdapter for non-empty annotations."""
    if annotation is inspect.Parameter.empty or annotation is Any:
        return None
    adapter = _TYPE_ADAPTER_CACHE.get(annotation)
    if adapter is None:
        try:
            adapter = TypeAdapter(annotation)
        except Exception:
            return None
        _TYPE_ADAPTER_CACHE[annotation] = adapter
    return adapter


@dataclass(frozen=True)
class ActionSpec:
    """Cached action metadata, alias mapping, and Pydantic validators."""

    name: str
    handler: _ToolFn
    description: str
    required: dict[str, str]
    optional: dict[str, dict[str, Any]]
    parameters: tuple[inspect.Parameter, ...]
    hints: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_handler(cls, name: str, handler: _ToolFn) -> ActionSpec:
        """Introspect signature, resolved type hints, and docstring summary once."""
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
        return cls(
            name=name,
            handler=handler,
            description=summary,
            required=required,
            optional=optional,
            parameters=tuple(sig.parameters.values()),
            hints=hints,
        )

    def to_help_dict(self) -> dict[str, Any]:
        """Render complete schema metadata for single-round-trip discovery."""
        return {
            "description": self.description,
            "required": dict(self.required),
            "optional": dict(self.optional),
        }

    def bind_and_validate(
        self, group: str, raw_params: dict[str, Any] | None
    ) -> dict[str, Any]:
        """Normalize aliases, reject unknown/missing params, and validate/coerce types."""
        where = f"{group} action '{self.name}'"
        params = dict(raw_params or {})
        valid_list = [p.name for p in self.parameters]
        valid_names = set(valid_list)

        for alias, canonical in _PARAM_ALIASES.items():
            if (
                alias in params
                and canonical in valid_names
                and alias not in valid_names
            ):
                if canonical in params:
                    raise ToolError(
                        f"{where}: provide either '{canonical}' or '{alias}', not both."
                    )
                params[canonical] = params.pop(alias)

        unknown = sorted(set(params) - valid_names)
        if unknown:
            hints = "".join(_format_suggestion(key, valid_list) for key in unknown)
            valid_str = ", ".join(valid_list) or "none"
            raise ToolError(
                f"{where}: unknown parameter(s) {unknown}.{hints} "
                f"Valid parameters: {valid_str}."
            )

        for req_name in self.required:
            if req_name not in params:
                req_str = (
                    ", ".join(f"{k} ({tp})" for k, tp in self.required.items())
                    or "none"
                )
                opt_str = ", ".join(self.optional) or "none"
                raise ToolError(
                    f"{where}: missing required parameter '{req_name}'. "
                    f"Required parameters: {req_str}. Optional parameters: {opt_str}."
                )

        kwargs: dict[str, Any] = {}
        for param in self.parameters:
            if param.name not in params:
                continue
            value = params[param.name]
            ann = self.hints.get(param.name, param.annotation)
            if _has_disallowed_bool(ann, value):
                expected = _format_type_name(ann)
                raise ToolError(
                    f"{where}: Invalid type for '{param.name}': "
                    f"expected {expected}, got bool"
                )
            adapter = _get_type_adapter(ann)
            if adapter is not None:
                try:
                    value = adapter.validate_python(value)
                except PydanticValidationError as exc:
                    expected = _format_type_name(ann)
                    raise ToolError(
                        f"{where}: Invalid type for '{param.name}': "
                        f"expected {expected}, got {type(value).__name__}"
                    ) from exc
            kwargs[param.name] = value

        return kwargs


@dataclass(frozen=True)
class CompoundGroupSpec:
    """Declarative specification for one consolidated compound tool."""

    name: str
    summary: str
    annotations: ToolAnnotations
    actions: dict[str, _ToolFn]


_ACTION_SPEC_CACHE: dict[tuple[str, _ToolFn], ActionSpec] = {}


def _get_action_spec(action: str, handler: _ToolFn) -> ActionSpec:
    """Return cached ActionSpec for (action, handler)."""
    key = (action, handler)
    spec = _ACTION_SPEC_CACHE.get(key)
    if spec is None:
        spec = ActionSpec.from_handler(action, handler)
        _ACTION_SPEC_CACHE[key] = spec
    return spec


def _require_params_dict(
    group: str, action: str, raw_params: Any
) -> dict[str, Any] | None:
    """Validate that optional `params` is a mapping before dispatching."""
    if raw_params is not None and not isinstance(raw_params, dict):
        raise ToolError(
            f"{group} action '{action}': `params` must be an object/dict, "
            f"got {type(raw_params).__name__}"
        )
    return raw_params


def _help_actions_payload(
    group: str,
    actions: dict[str, _ToolFn],
    names: list[str],
    *,
    include_hint: bool = False,
) -> dict[str, Any]:
    """Assemble help schema payload for a list of action names."""
    payload: dict[str, Any] = {
        "tool": group,
        "actions": {
            name: _get_action_spec(name, actions[name]).to_help_dict() for name in names
        },
        "aliases": dict(_PARAM_ALIASES),
    }
    if include_hint:
        payload["hint"] = (
            "Pass action='<name>' with the required params above, "
            "or filter with params={'action': '<name>'}."
        )
    return payload


def _handle_help(
    group: str, actions: dict[str, _ToolFn], params: dict[str, Any] | None
) -> dict[str, Any]:
    """Return full parameter schemas for all actions in one call, or inspect one action."""
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
        return {
            "tool": group,
            "action": target,
            **_get_action_spec(target, actions[target]).to_help_dict(),
        }

    if params and "actions" in params:
        subset = params["actions"]
        if (
            not isinstance(subset, list)
            or not subset
            or not all(isinstance(name, str) and name in actions for name in subset)
        ):
            valid = ", ".join(sorted(actions))
            raise ToolError(
                f"{group} action 'help': `actions` must be a non-empty list of valid action names ({valid})."
            )
        return _help_actions_payload(group, actions, subset)

    return _help_actions_payload(group, actions, sorted(actions), include_hint=True)


async def _dispatch(
    group: str,
    actions: dict[str, _ToolFn],
    action: str,
    params: dict[str, Any] | None,
) -> dict[str, Any]:
    """Validate action/params against cached ActionSpec and invoke the handler."""
    validated_params = _require_params_dict(group, action, params)
    if action == "help":
        return _handle_help(group, actions, validated_params)

    handler = actions.get(action)
    if handler is None:
        valid = ", ".join(sorted(actions))
        suggestion = _format_suggestion(action, sorted(actions))
        raise ToolError(
            f"Unknown action '{action}' for {group}.{suggestion} "
            f"Valid: {valid}. Call with action='help' for parameter details."
        )

    spec = _get_action_spec(action, handler)
    kwargs = spec.bind_and_validate(group, validated_params)
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

_WRITE_ANNOTATIONS = ToolAnnotations(read_only_hint=False)
_DESTRUCTIVE_ANNOTATIONS = ToolAnnotations(read_only_hint=False, destructive_hint=True)

_GROUP_SPECS: tuple[CompoundGroupSpec, ...] = (
    CompoundGroupSpec(
        name="word_document",
        summary="Document lifecycle, inspection, search, merge, and PDF export.",
        annotations=_WRITE_ANNOTATIONS,
        actions=_DOCUMENT_ACTIONS,
    ),
    CompoundGroupSpec(
        name="word_content",
        summary="Add, update, search/replace, or delete document content.",
        annotations=_DESTRUCTIVE_ANNOTATIONS,
        actions=_CONTENT_ACTIONS,
    ),
    CompoundGroupSpec(
        name="word_format",
        summary="Format text, tables, styles, headers, and footers.",
        annotations=_WRITE_ANNOTATIONS,
        actions=_FORMAT_ACTIONS,
    ),
    CompoundGroupSpec(
        name="word_protection",
        summary="Encrypt, unprotect, restrict editing, add signature marker, or verify.",
        annotations=_WRITE_ANNOTATIONS,
        actions=_PROTECTION_ACTIONS,
    ),
    CompoundGroupSpec(
        name="word_footnote",
        summary="Add footnotes/endnotes, convert footnotes, or style notes.",
        annotations=_WRITE_ANNOTATIONS,
        actions=_FOOTNOTE_ACTIONS,
    ),
    CompoundGroupSpec(
        name="word_property",
        summary="Read/update core document properties and page layout.",
        annotations=_WRITE_ANNOTATIONS,
        actions=_PROPERTY_ACTIONS,
    ),
)


def _build_group_description(spec: CompoundGroupSpec) -> str:
    """Build concise description listing actions without repeating boilerplate."""
    action_list = ", ".join(spec.actions)
    return (
        f"{spec.summary} Actions: {action_list}. "
        "Pass args in `params`; use action='help' for schemas."
    )


def _make_group_tool(spec: CompoundGroupSpec) -> _ToolFn:
    """Build a bound FastMCP dispatcher coroutine for one compound domain."""

    async def _group_tool(
        action: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        return await _dispatch(spec.name, spec.actions, action, params)

    _group_tool.__name__ = spec.name
    return _group_tool


def register_compound_tools(mcp: FastMCP) -> None:
    """Register the 9-tool hybrid compound surface on the FastMCP server."""
    mcp.tool(annotations=_WRITE_ANNOTATIONS)(batch_tools.build_document)
    mcp.tool(annotations=_WRITE_ANNOTATIONS)(batch_tools.create_from_markdown)
    mcp.tool(annotations=_WRITE_ANNOTATIONS)(batch_tools.batch_add_content)

    for group_spec in _GROUP_SPECS:
        mcp.tool(
            name=group_spec.name,
            annotations=group_spec.annotations,
            description=_build_group_description(group_spec),
        )(_make_group_tool(group_spec))
