"""Exception handling utilities for MCP Word Server.

This module provides tools and utilities for consistent exception handling
across all MCP tools, including error mapping and canonical OperationError builders.
"""

from __future__ import annotations

import logging
import sys
from typing import Any

from docx.opc.exceptions import PackageNotFoundError

from ..models.response_models import OperationError
from .exception_core import (
    ConfigurationError,
    DocumentProcessingError,
    FileOperationError,
    SecurityError,
    StyleError,
    ValidationError,
)


logger = logging.getLogger(__name__)


class ExceptionTool:
    """Single authoritative translator from exceptions into OperationError payloads."""

    # Specific subclasses must precede base classes (e.g. SecurityError before PermissionError before OSError).
    EXCEPTION_MAPPING: dict[type[Exception], dict[str, Any]] = {
        SecurityError: {
            "error_type": "security_violation",
            "recoverable": False,
        },
        DocumentProcessingError: {
            "error_type": "document_processing",
            "recoverable": True,
        },
        ValidationError: {
            "error_type": "validation",
            "recoverable": False,
        },
        FileOperationError: {
            "error_type": "file_operation",
            "recoverable": True,
        },
        StyleError: {
            "error_type": "style",
            "recoverable": True,
        },
        ConfigurationError: {
            "error_type": "configuration",
            "recoverable": False,
        },
        ValueError: {
            "error_type": "validation",
            "recoverable": False,
        },
        IndexError: {
            "error_type": "validation",
            "recoverable": False,
        },
        KeyError: {
            "error_type": "validation",
            "recoverable": False,
        },
        PackageNotFoundError: {
            "error_type": "document_processing",
            "recoverable": True,
        },
        FileNotFoundError: {
            "error_type": "file_not_found",
            "recoverable": False,
        },
        PermissionError: {
            "error_type": "permission_denied",
            "recoverable": False,
        },
        IOError: {
            "error_type": "file_operation",
            "recoverable": True,
        },
        OSError: {
            "error_type": "file_operation",
            "recoverable": True,
        },
    }

    @staticmethod
    def get_error_info(exception: Exception) -> dict[str, Any]:
        """Get error information for an exception.

        Args:
            exception: The exception to analyze.

        Returns:
            Dictionary with error_type, recoverable, and message.
        """
        exc_type = type(exception)

        for exc_class, info in ExceptionTool.EXCEPTION_MAPPING.items():
            if issubclass(exc_type, exc_class):
                return {
                    "error_type": info["error_type"],
                    "recoverable": info["recoverable"],
                    "message": str(exception),
                }

        return {
            "error_type": "unknown",
            "recoverable": True,
            "message": str(exception),
        }

    @staticmethod
    def to_operation_error(
        exception: Exception,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> OperationError:
        """Convert an exception to an OperationError response.

        Suggestion precedence:
            1. Explicit ``suggestion`` argument
            2. ``getattr(exception, "suggestion", None)``
            3. ``_get_suggestion(error_type, exception)``
        """
        error_info = ExceptionTool.get_error_info(exception)
        exc_suggestion = getattr(exception, "suggestion", None)
        effective_suggestion = (
            suggestion
            if suggestion is not None
            else (
                exc_suggestion
                if exc_suggestion is not None
                else ExceptionTool._get_suggestion(error_info["error_type"], exception)
            )
        )

        merged_details: dict[str, Any] = {}
        exc_details = getattr(exception, "details", None)
        if isinstance(exc_details, dict):
            merged_details.update(exc_details)
        if details:
            merged_details.update(details)

        return OperationError(
            status="error",
            error_type=error_info["error_type"],
            message=error_info["message"],
            suggestion=effective_suggestion,
            recoverable=error_info["recoverable"],
            details=merged_details or None,
        )

    @staticmethod
    def handle_error(
        exception: Exception,
        filename: str | None = None,
        operation: str | None = None,
        suggestion: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Handle an exception and return a standardized OperationError dictionary."""
        op_error = ExceptionTool.to_operation_error(
            exception, suggestion=suggestion, details=details
        )

        context: list[str] = []
        if filename:
            context.append(f"File: {filename}")
        if operation:
            context.append(f"Operation: {operation}")

        message = op_error.message
        if context:
            message = f"{message} ({', '.join(context)})"

        from ..utils.logging_config import is_dev_mode

        if is_dev_mode():
            exc_info: Exception | bool = (
                exception
                if exception.__traceback__ is not None
                else sys.exc_info()[2] is not None
            )
            logger.error("%s", message, exc_info=exc_info)

        final_error = op_error.model_copy(update={"message": message})
        exclude = {"details"} if final_error.details is None else set()
        return final_error.model_dump(exclude=exclude)

    @staticmethod
    def _get_suggestion(error_type: str, exception: Exception) -> str | None:
        """Get a helpful suggestion based on error type."""
        suggestions = {
            "security_violation": (
                "Move the target file into an allowed directory or configure "
                "WORD_MCP_DIRS to include its directory."
            ),
            "document_processing": "Check if the document is corrupted or in a valid format.",
            "validation": "Review the input parameters and ensure they meet the requirements.",
            "file_operation": "Verify file permissions and that the file is not open in another program.",
            "file_not_found": "Check that the file path is correct and the file exists.",
            "permission_denied": "Ensure you have the necessary permissions to access the file.",
            "style": "Verify the style name exists in the document template.",
            "configuration": "Check the application configuration settings.",
        }
        return suggestions.get(error_type)
