"""Runtime logging and environment mode configuration for Word MCP Server."""

from __future__ import annotations

import logging
import os
import sys


VALID_LOG_LEVELS = ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL")
_LOG_LEVEL_ALIASES = {"WARN": "WARNING"}


def is_dev_mode() -> bool:
    """Return True when WORD_MCP_DEV_MODE is enabled."""
    raw = (os.environ.get("WORD_MCP_DEV_MODE") or "false").strip().lower()
    return raw in ("1", "true", "yes", "on")


def resolve_log_level(explicit_level: str | None = None) -> int:
    """Resolve the active logging level from explicit override or WORD_MCP_LOG_LEVEL."""
    raw = (
        (
            explicit_level
            or (os.environ.get("WORD_MCP_LOG_LEVEL") or "").strip()
            or ("DEBUG" if is_dev_mode() else "INFO")
        )
        .strip()
        .upper()
    )
    normalized = _LOG_LEVEL_ALIASES.get(raw, raw)
    if normalized not in VALID_LOG_LEVELS:
        valid = ", ".join(VALID_LOG_LEVELS)
        raise ValueError(
            f"Invalid WORD_MCP_LOG_LEVEL '{raw}'. Expected one of: {valid}."
        )
    return int(getattr(logging, normalized))


def configure_logging(log_level: str | None = None) -> logging.Logger:
    """Configure the word_mcp logger to write exclusively to sys.stderr."""
    level = resolve_log_level(log_level)
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    )
    root_logger = logging.getLogger("word_mcp")
    root_logger.handlers.clear()
    root_logger.addHandler(handler)
    root_logger.setLevel(level)
    root_logger.propagate = False
    return root_logger
