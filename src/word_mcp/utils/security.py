"""Security utilities for file path validation.

This module provides functions for validating file paths against
allowed directories for security purposes.
"""

from __future__ import annotations

import os
from pathlib import Path

from word_mcp.exception.exception_core import SecurityError, ValidationError


def get_allowed_directories() -> list[str]:
    """Get the list of allowed directories from environment variables.

    Reads ``WORD_MCP_DIRS`` (falling back to legacy ``MCP_ALLOWED_DIRECTORIES``).

    Returns:
        List of absolute paths to directories where documents can be created/accessed.
        Defaults to ``['./documents']`` if neither variable is set.
    """
    allowed_dirs_str = (
        os.environ.get("WORD_MCP_DIRS")
        or os.environ.get("MCP_ALLOWED_DIRECTORIES")
        or "./documents"
    )
    allowed_dirs = [d.strip() for d in allowed_dirs_str.split(",") if d.strip()]
    return [str(Path(d).resolve()) for d in allowed_dirs]


def resolve_sandboxed_path(
    raw_path: str | Path,
    *,
    param_name: str = "filename",
) -> Path:
    """Validate path syntax and enforce WORD_MCP_DIRS containment.

    Raises:
        ValidationError: If the path is empty, whitespace-only, contains NUL, or is invalid.
        SecurityError: If the resolved path falls outside WORD_MCP_DIRS.
    """
    raw = os.fspath(raw_path)
    if not raw.strip():
        raise ValidationError(
            f"Parameter '{param_name}' must be a non-empty valid path.",
            suggestion=f"Provide a non-empty file path for '{param_name}'.",
            details={"param": param_name},
        )
    if "\x00" in raw:
        raise ValidationError(
            f"Parameter '{param_name}' contains an invalid NUL byte.",
            suggestion="Remove control characters from the file path.",
            details={"param": param_name},
        )

    allowed_dirs = get_allowed_directories()
    if not allowed_dirs:
        raise SecurityError(
            "No allowed directories are configured in WORD_MCP_DIRS.",
            suggestion="Configure WORD_MCP_DIRS with one or more allowed directories.",
            details={"param": param_name, "path": raw},
        )

    try:
        resolved = Path(raw).resolve()
    except (OSError, ValueError) as exc:
        raise ValidationError(
            f"Invalid path for '{param_name}': {exc}",
            suggestion=f"Provide a valid path for '{param_name}'.",
            details={"param": param_name, "path": raw},
        ) from exc

    for allowed_dir in allowed_dirs:
        allowed_path = Path(allowed_dir).resolve()
        try:
            if resolved == allowed_path or resolved.is_relative_to(allowed_path):
                return resolved
        except (ValueError, TypeError):
            continue

    raise SecurityError(
        f"Path '{raw}' is not in allowed directories: {', '.join(allowed_dirs)}",
        suggestion=(
            "Move the target into an allowed directory or configure "
            "WORD_MCP_DIRS to include its directory."
        ),
        details={"param": param_name, "path": raw},
    )


def is_path_in_allowed_directories(file_path: str | Path) -> tuple[bool, str | None]:
    """Check if the given file path is within allowed directories."""
    try:
        resolve_sandboxed_path(file_path)
        return True, None
    except (ValidationError, SecurityError) as exc:
        return False, str(exc)
