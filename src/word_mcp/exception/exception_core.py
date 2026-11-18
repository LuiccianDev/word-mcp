"""Core exceptions for MCP Word Server.

This module defines all custom exceptions used throughout the application
for better error handling and more informative error messages.
"""

from __future__ import annotations

from typing import Any


class _ContextualErrorMixin:
    """Attach optional actionable recovery metadata to domain exceptions."""

    suggestion: str | None
    details: dict[str, Any] | None

    def _init_context(
        self,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        self.suggestion = suggestion
        self.details = details


class DocumentProcessingError(_ContextualErrorMixin, Exception):
    """Raised when there is an error processing a Word document.

    This exception should be used for errors that occur during document
    manipulation, such as adding content, formatting, or saving documents.
    """

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)


class ValidationError(_ContextualErrorMixin, ValueError):
    """Raised when input validation fails.

    This exception should be used for errors related to invalid input
    parameters or data that doesn't meet the required criteria.
    """

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)


class FileOperationError(_ContextualErrorMixin, IOError):
    """Raised when there is an error performing file operations.

    This includes errors related to reading, writing, or accessing files.
    """

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)


class SecurityError(_ContextualErrorMixin, PermissionError):
    """Raised when a path escapes configured WORD_MCP_DIRS boundaries.

    Inherits from PermissionError so OS permission expectations remain valid
    while mapping explicitly to 'security_violation' in ExceptionTool.
    """

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)


class StyleError(_ContextualErrorMixin, Exception):
    """Raised when there is an error related to document styles.

    This includes missing styles, invalid style names, or style application errors.
    """

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)


class ConfigurationError(_ContextualErrorMixin, Exception):
    """Raised when there is a configuration error in the application.

    This includes missing or invalid configuration settings.
    """

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)


class ContextualFileNotFoundError(_ContextualErrorMixin, FileNotFoundError):
    """FileNotFoundError carrying an actionable recovery suggestion."""

    def __init__(
        self,
        message: str,
        *,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self._init_context(suggestion=suggestion, details=details)
