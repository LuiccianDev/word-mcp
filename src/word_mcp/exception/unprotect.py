"""Compatibility re-export for removing document protection."""

from __future__ import annotations


def remove_protection_info(
    filename: str, password: str | None = None
) -> tuple[bool, str]:
    """Remove protection metadata and decrypt the document if encrypted."""
    from word_mcp.core.protection import remove_protection_info as _remove

    return _remove(filename, password)
