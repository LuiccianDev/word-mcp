"""
Validation functions for the Word Document Server.

This package contains declarative boundary rules and validators for document tools.
"""

from word_mcp.validation.document_validators import (
    PathAccess,
    PathRule,
    tool_boundary,
    validate_docx_create,
    validate_docx_read,
    validate_docx_write,
    validate_file_read,
    validate_file_write,
    validate_path,
)


__all__ = [
    "PathAccess",
    "PathRule",
    "tool_boundary",
    "validate_docx_create",
    "validate_docx_read",
    "validate_docx_write",
    "validate_file_read",
    "validate_file_write",
    "validate_path",
]
