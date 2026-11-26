"""Declarative validation boundary for Word document tool parameters.

Provides sandbox-first path validation, lightweight OOXML ZIP preflight checks,
and a unified OperationError boundary decorator for all MCP tools.
"""

from __future__ import annotations

import functools
import inspect
import os
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal, TypeVar, cast

from fastmcp.exceptions import ToolError

from word_mcp.exception.exception_core import (
    ContextualFileNotFoundError,
    DocumentProcessingError,
    FileOperationError,
    ValidationError,
)
from word_mcp.exception.exception_tools import ExceptionTool
from word_mcp.utils.security import resolve_sandboxed_path


F = TypeVar("F", bound=Callable[..., Any])

MAX_ZIP_ENTRIES = 5_000
MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024
MAX_SINGLE_ENTRY_BYTES = 50 * 1024 * 1024
MAX_COMPRESSION_RATIO = 100
RATIO_CHECK_THRESHOLD_BYTES = 1024 * 1024


class PathAccess(StrEnum):
    """Access policies enforced by the tool boundary before executing tool bodies."""

    DOCX_READ = "docx_read"
    DOCX_WRITE = "docx_write"
    DOCX_CREATE = "docx_create"
    FILE_READ_SHALLOW = "file_read_shallow"
    FILE_WRITE_SHALLOW = "file_write_shallow"
    SANDBOX_FILE_READ = "sandbox_file_read"
    SANDBOX_FILE_WRITE = "sandbox_file_write"
    SANDBOX_DIR_READ = "sandbox_dir_read"


@dataclass(frozen=True)
class _AccessPolicy:
    kind: Literal["create_file", "read_dir", "read_file"]
    require_docx: bool = False
    check_ooxml: bool = False
    require_writeable: bool = False


_ACCESS_POLICIES: dict[PathAccess, _AccessPolicy] = {
    PathAccess.DOCX_READ: _AccessPolicy(
        kind="read_file", require_docx=True, check_ooxml=True
    ),
    PathAccess.DOCX_WRITE: _AccessPolicy(
        kind="read_file", require_docx=True, check_ooxml=True, require_writeable=True
    ),
    PathAccess.DOCX_CREATE: _AccessPolicy(
        kind="create_file", require_docx=True, require_writeable=True
    ),
    PathAccess.FILE_READ_SHALLOW: _AccessPolicy(kind="read_file", require_docx=True),
    PathAccess.FILE_WRITE_SHALLOW: _AccessPolicy(
        kind="read_file", require_docx=True, require_writeable=True
    ),
    PathAccess.SANDBOX_FILE_READ: _AccessPolicy(kind="read_file"),
    PathAccess.SANDBOX_FILE_WRITE: _AccessPolicy(
        kind="create_file", require_writeable=True
    ),
    PathAccess.SANDBOX_DIR_READ: _AccessPolicy(kind="read_dir"),
}


@dataclass(frozen=True)
class PathRule:
    """Declarative validation contract for a single tool path parameter."""

    access: PathAccess
    required: bool = True
    many: bool = False
    aliases: tuple[str, ...] = ()
    default_ext: str | None = None
    fallback_from: str | None = None
    fallback_suffix: str | None = None


def _fast_check_ooxml_zip(path: Path) -> None:
    """Verify OOXML ZIP container structure and decompression bounds."""
    try:
        # Open ZipFile directly instead of zipfile.is_zipfile so OS sharing/lock
        # PermissionErrors surface as file_operation rather than looking corrupt.
        with zipfile.ZipFile(path, "r") as zf:
            infos = zf.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                raise DocumentProcessingError(
                    f"Archive '{path}' exceeds maximum entry count ({len(infos)} > {MAX_ZIP_ENTRIES}).",
                    suggestion="Reduce the number of embedded objects in the .docx archive.",
                )
            total_uncompressed = 0
            for info in infos:
                if info.file_size > MAX_SINGLE_ENTRY_BYTES:
                    raise DocumentProcessingError(
                        f"Archive entry '{info.filename}' exceeds maximum uncompressed size ({info.file_size} > {MAX_SINGLE_ENTRY_BYTES}).",
                        suggestion="The document contains an oversized embedded part (potential ZIP bomb).",
                    )
                total_uncompressed += info.file_size
                if total_uncompressed > MAX_UNCOMPRESSED_BYTES:
                    raise DocumentProcessingError(
                        f"Archive '{path}' exceeds maximum total uncompressed size ({total_uncompressed} > {MAX_UNCOMPRESSED_BYTES}).",
                        suggestion="The document exceeds safe uncompressed size limits (potential ZIP bomb).",
                    )
                if info.file_size >= RATIO_CHECK_THRESHOLD_BYTES:
                    compressed = max(1, info.compress_size)
                    if info.file_size / compressed > MAX_COMPRESSION_RATIO:
                        raise DocumentProcessingError(
                            f"Archive entry '{info.filename}' exceeds maximum compression ratio.",
                            suggestion="The document appears to contain an abnormal compression ratio (potential ZIP bomb).",
                        )
            names = {info.filename for info in infos}
    except zipfile.BadZipFile as exc:
        raise DocumentProcessingError(
            f"File '{path}' is not a valid Word document (.docx).",
            suggestion=(
                "The file may be corrupt, encrypted, or a non-.docx format renamed "
                "with a .docx extension. If encrypted, unprotect it first."
            ),
        ) from exc
    except OSError as exc:
        raise FileOperationError(
            f"Could not open document '{path}': {exc}",
            suggestion=(
                "Ensure the file is not locked by Microsoft Word or another process "
                "and that you have read permissions."
            ),
        ) from exc

    if "word/document.xml" not in names or "[Content_Types].xml" not in names:
        raise DocumentProcessingError(
            f"File '{path}' is not a valid Word document (.docx).",
            suggestion=(
                "Missing required OOXML parts ('word/document.xml' or "
                "'[Content_Types].xml') inside the .docx archive."
            ),
        )


def _check_file_writeable(path_value: str | Path) -> tuple[bool, str]:
    """Check if a file is writeable or can be created in its parent directory."""
    try:
        path = Path(path_value).resolve()
        parent = path.parent
        if not parent.exists():
            return False, f"Parent directory '{parent}' does not exist."
        if not os.access(parent, os.W_OK):
            return False, f"Parent directory '{parent}' is not writeable."
        if path.exists() and not os.access(path, os.W_OK):
            return False, f"File '{path}' is not writeable."
        return True, ""
    except Exception as exc:
        return False, str(exc)


def _normalize_extension(raw_path: str, default_ext: str | None) -> str:
    if not default_ext or not raw_path.strip() or "\x00" in raw_path:
        return raw_path
    ext = default_ext if default_ext.startswith(".") else f".{default_ext}"
    if not raw_path.lower().endswith(ext.lower()):
        return f"{raw_path}{ext}"
    return raw_path


def _ensure_writeable(resolved: Path) -> None:
    ok, err = _check_file_writeable(resolved)
    if not ok:
        raise FileOperationError(
            f"Cannot write to file: {err}",
            suggestion=(
                "Ensure the file is not open in Microsoft Word, verify write permissions, "
                "and check that the target directory exists and is within WORD_MCP_DIRS."
            ),
        )


def _require_docx_suffix(resolved: Path, param_name: str, is_existing: bool) -> None:
    if resolved.suffix.lower() != ".docx":
        hint = (
            "Provide a path ending in '.docx'. Legacy '.doc' files must be "
            "converted to '.docx' before editing."
            if is_existing
            else "Provide a path ending in '.docx'."
        )
        raise ValidationError(
            f"File '{resolved}' is not a .docx document.",
            suggestion=hint,
            details={"param": param_name},
        )


def validate_path(
    raw_path: Any,
    access: PathAccess,
    *,
    param_name: str = "filename",
    default_ext: str | None = None,
) -> Path:
    """Enforce sandbox containment, file/dir preconditions, and fast OOXML checks."""
    if not isinstance(raw_path, str):
        raise ValidationError(
            f"Parameter '{param_name}' must be a string path.",
            suggestion=f"Provide a valid file path string for '{param_name}'.",
            details={"param": param_name},
        )

    candidate = _normalize_extension(raw_path, default_ext)
    # Sandbox check runs first so out-of-sandbox paths never leak filesystem existence.
    resolved = resolve_sandboxed_path(candidate, param_name=param_name)
    policy = _ACCESS_POLICIES[access]

    if policy.kind == "create_file":
        if policy.require_docx:
            _require_docx_suffix(resolved, param_name, is_existing=False)
        if resolved.exists() and resolved.is_dir():
            raise ValidationError(
                f"Path '{resolved}' is a directory.",
                suggestion=f"Provide a file path rather than a directory for '{param_name}'.",
                details={"param": param_name},
            )
        resolved.parent.mkdir(parents=True, exist_ok=True)
        if policy.require_writeable:
            _ensure_writeable(resolved)
        return resolved

    if policy.kind == "read_dir":
        if not resolved.exists():
            raise ContextualFileNotFoundError(
                f"Directory '{resolved}' does not exist.",
                suggestion="Verify the directory path exists within WORD_MCP_DIRS.",
            )
        if not resolved.is_dir():
            raise ValidationError(
                f"Path '{resolved}' is not a directory.",
                suggestion="Provide an existing directory path within WORD_MCP_DIRS.",
            )
        return resolved

    if not resolved.exists():
        suggestion = (
            "Verify the path with list_available_documents, or create the document "
            "first using create_document or build_document."
            if policy.require_docx
            else "Verify the file path exists within WORD_MCP_DIRS."
        )
        raise ContextualFileNotFoundError(
            f"File '{resolved}' does not exist.",
            suggestion=suggestion,
        )

    if not resolved.is_file():
        raise ValidationError(
            f"Path '{resolved}' is not a regular file.",
            suggestion=f"Provide a valid file path within WORD_MCP_DIRS for '{param_name}'.",
            details={"param": param_name},
        )

    if policy.require_docx:
        _require_docx_suffix(resolved, param_name, is_existing=True)

    if policy.check_ooxml:
        _fast_check_ooxml_zip(resolved)

    if policy.require_writeable:
        _ensure_writeable(resolved)

    return resolved


def _validate_docx_path(path_str: str) -> dict[str, Any] | None:
    """Validate DOCX_READ preconditions; returns canonical OperationError dict or None."""
    try:
        validate_path(path_str, PathAccess.DOCX_READ, param_name="filename")
        return None
    except Exception as exc:
        return ExceptionTool.handle_error(
            exc, filename=path_str, operation="validate document"
        )


def _extract_primary_hint(
    sig: inspect.Signature,
    primary_path_param: str,
    rules: dict[str, PathRule],
    args: tuple[Any, ...],
    kwargs: dict[str, Any],
) -> str | None:
    """Extract best-effort filename for error context before validation raises."""
    candidates = [primary_path_param]
    if primary_path_param in rules:
        candidates.extend(rules[primary_path_param].aliases)
    param_names = list(sig.parameters.keys())
    for name in candidates:
        if name in kwargs:
            kw_val = kwargs[name]
            if isinstance(kw_val, str) and kw_val:
                return kw_val
        if name in param_names:
            idx = param_names.index(name)
            if idx < len(args):
                arg_val = args[idx]
                if isinstance(arg_val, str) and arg_val:
                    return arg_val
    return None


def _resolve_rule_value(
    param_name: str,
    rule: PathRule,
    arguments: dict[str, Any],
    extra_kwargs: dict[str, Any],
) -> Any:
    """Resolve parameter value from canonical param, aliases, or fallback."""
    primary_val = arguments.get(param_name)
    alias_hits: list[tuple[str, Any]] = []
    for alias in rule.aliases:
        if alias in arguments and arguments[alias] is not None:
            alias_hits.append((alias, arguments[alias]))
        elif alias in extra_kwargs and extra_kwargs[alias] is not None:
            alias_hits.append((alias, extra_kwargs[alias]))

    for alias_name, alias_val in alias_hits:
        if primary_val is not None and primary_val != "" and alias_val != "":
            if str(primary_val) != str(alias_val):
                raise ValidationError(
                    f"Provide either '{param_name}' or '{alias_name}', not both with different values.",
                    suggestion=f"Use '{param_name}' as the canonical path parameter.",
                )
        elif primary_val is None or primary_val == "":
            primary_val = alias_val

    if primary_val is None and rule.fallback_from:
        fallback_val = arguments.get(rule.fallback_from)
        if isinstance(fallback_val, str) and fallback_val.strip():
            primary_val = str(
                Path(fallback_val).with_suffix(rule.fallback_suffix or "")
            )

    return primary_val


def tool_boundary(
    rules: dict[str, PathRule],
    *,
    operation: str | None = None,
    primary_path_param: str = "filename",
) -> Callable[[F], F]:
    """Unified preflight, sandbox, normalization, and OperationError boundary decorator."""

    def decorator(func: F) -> F:
        sig = inspect.signature(func)
        op_name = operation or func.__name__.replace("_", " ")
        is_coro = inspect.iscoroutinefunction(func)

        def _prepare_call(
            args: tuple[Any, ...], kwargs: dict[str, Any]
        ) -> tuple[tuple[Any, ...], dict[str, Any], str | None]:
            call_kwargs = dict(kwargs)
            extra_kwargs: dict[str, Any] = {}
            for rule in rules.values():
                for alias in rule.aliases:
                    if alias in call_kwargs and alias not in sig.parameters:
                        extra_kwargs[alias] = call_kwargs.pop(alias)

            bound = sig.bind_partial(*args, **call_kwargs)
            bound.apply_defaults()
            primary_hint = _extract_primary_hint(
                sig, primary_path_param, rules, args, kwargs
            )

            for param_name, rule in rules.items():
                raw_val = _resolve_rule_value(
                    param_name, rule, bound.arguments, extra_kwargs
                )
                if raw_val is None:
                    if rule.required:
                        raise ValidationError(
                            f"Missing required parameter '{param_name}'.",
                            suggestion=f"Provide a valid path for '{param_name}'.",
                            details={"param": param_name},
                        )
                    continue

                if rule.many:
                    if not isinstance(raw_val, list) or not raw_val:
                        raise ValidationError(
                            f"Parameter '{param_name}' must be a non-empty list of file paths.",
                            suggestion=f"Provide one or more valid paths in '{param_name}'.",
                            details={"param": param_name},
                        )
                    bound.arguments[param_name] = [
                        str(
                            validate_path(
                                item,
                                rule.access,
                                param_name=param_name,
                                default_ext=rule.default_ext,
                            )
                        )
                        for item in raw_val
                    ]
                else:
                    resolved = validate_path(
                        raw_val,
                        rule.access,
                        param_name=param_name,
                        default_ext=rule.default_ext,
                    )
                    if param_name in bound.arguments:
                        bound.arguments[param_name] = str(resolved)
                    for alias in rule.aliases:
                        if alias in bound.arguments:
                            bound.arguments[alias] = None
                    if param_name == primary_path_param:
                        primary_hint = str(resolved)

            return bound.args, bound.kwargs, primary_hint

        def _on_error(
            exc: Exception,
            primary: str | None,
            args: tuple[Any, ...],
            kwargs: dict[str, Any],
        ) -> dict[str, Any]:
            if primary is None:
                primary = _extract_primary_hint(
                    sig, primary_path_param, rules, args, kwargs
                )
            return ExceptionTool.handle_error(exc, filename=primary, operation=op_name)

        if is_coro:

            @functools.wraps(func)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                primary: str | None = None
                try:
                    call_args, call_kwargs, primary = _prepare_call(args, kwargs)
                    return await func(*call_args, **call_kwargs)
                except ToolError:
                    raise
                except Exception as exc:
                    return _on_error(exc, primary, args, kwargs)

            return cast(F, async_wrapper)

        @functools.wraps(func)
        def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
            primary: str | None = None
            try:
                call_args, call_kwargs, primary = _prepare_call(args, kwargs)
                return func(*call_args, **call_kwargs)
            except ToolError:
                raise
            except Exception as exc:
                return _on_error(exc, primary, args, kwargs)

        return cast(F, sync_wrapper)

    return decorator


def validate_docx_read(
    param_name: str, *, operation: str | None = None
) -> Callable[[F], F]:
    """Boundary preset for read-only .docx tools."""
    return tool_boundary(
        {param_name: PathRule(PathAccess.DOCX_READ)},
        operation=operation,
        primary_path_param=param_name,
    )


def validate_docx_write(
    param_name: str, *, operation: str | None = None
) -> Callable[[F], F]:
    """Boundary preset for read-write .docx tools."""
    return tool_boundary(
        {param_name: PathRule(PathAccess.DOCX_WRITE)},
        operation=operation,
        primary_path_param=param_name,
    )


def validate_docx_create(
    param_name: str, *, operation: str | None = None
) -> Callable[[F], F]:
    """Boundary preset for creating or overwriting a .docx file."""
    return tool_boundary(
        {param_name: PathRule(PathAccess.DOCX_CREATE, default_ext=".docx")},
        operation=operation,
        primary_path_param=param_name,
    )


def validate_file_read(
    param_name: str, *, operation: str | None = None
) -> Callable[[F], F]:
    """Boundary preset for shallow read of potentially encrypted .docx files."""
    return tool_boundary(
        {param_name: PathRule(PathAccess.FILE_READ_SHALLOW)},
        operation=operation,
        primary_path_param=param_name,
    )


def validate_file_write(
    param_name: str, *, operation: str | None = None
) -> Callable[[F], F]:
    """Boundary preset for shallow write of potentially encrypted .docx files."""
    return tool_boundary(
        {param_name: PathRule(PathAccess.FILE_WRITE_SHALLOW)},
        operation=operation,
        primary_path_param=param_name,
    )
