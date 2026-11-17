"""Runtime logging and environment mode configuration for Word MCP Server."""

from __future__ import annotations

import logging
import os
import sys

VALID_LOG_LEVELS = ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL")


def is_dev_mode() -> bool:
    """Return True when WORD_MCP_DEV_MODE or DEV_MODE is enabled."""
    raw = os.environ.get("WORD_MCP_DEV_MODE") or os.environ.get("DEV_MODE") or "false"
    return raw.strip().lower() in ("1", "true", "yes", "on")


def resolve_log_level(explicit_level: str | None = None) -> int:
    """Resolve the active logging level from explicit override or environment."""
    raw = (
        explicit_level
        or os.environ.get("WORD_MCP_LOG_LEVEL")
        or os.environ.get("LOG_LEVEL")
        or ("DEBUG" if is_dev_mode() else "INFO")
    ).strip().upper()
    if raw not in VALID_LOG_LEVELS:
        valid = ", ".join(VALID_LOG_LEVELS)
        raise ValueError(f"Invalid LOG_LEVEL '{raw}'. Expected one of: {valid}.")
    return int(getattr(logging, raw))


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
