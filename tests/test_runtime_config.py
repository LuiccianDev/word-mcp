"""Behavioral tests for runtime logging and environment configuration."""

import io
import logging
import os
import unittest
from contextlib import redirect_stderr, redirect_stdout


class TestRuntimeConfig(unittest.TestCase):
    """Verify LOG_LEVEL and DEV_MODE behavior and stderr-only logging."""

    def setUp(self) -> None:
        self._env_backup = os.environ.copy()

    def tearDown(self) -> None:
        os.environ.clear()
        os.environ.update(self._env_backup)

    def test_configure_logging_writes_only_to_stderr(self) -> None:
        from word_mcp.utils.logging_config import configure_logging

        os.environ["LOG_LEVEL"] = "INFO"
        stdout_buf = io.StringIO()
        stderr_buf = io.StringIO()
        with redirect_stdout(stdout_buf), redirect_stderr(stderr_buf):
            logger = configure_logging()
            logger.info("startup check")

        self.assertEqual(stdout_buf.getvalue(), "")
        self.assertIn("startup check", stderr_buf.getvalue())
        self.assertEqual(logger.level, logging.INFO)

    def test_word_mcp_log_level_takes_precedence_over_log_level(self) -> None:
        from word_mcp.utils.logging_config import resolve_log_level

        os.environ["LOG_LEVEL"] = "WARNING"
        os.environ["WORD_MCP_LOG_LEVEL"] = "ERROR"
        self.assertEqual(resolve_log_level(), logging.ERROR)

    def test_dev_mode_defaults_log_level_to_debug(self) -> None:
        from word_mcp.utils.logging_config import is_dev_mode, resolve_log_level

        os.environ.pop("LOG_LEVEL", None)
        os.environ.pop("WORD_MCP_LOG_LEVEL", None)
        os.environ["DEV_MODE"] = "true"
        self.assertTrue(is_dev_mode())
        self.assertEqual(resolve_log_level(), logging.DEBUG)

    def test_invalid_log_level_raises_value_error(self) -> None:
        from word_mcp.utils.logging_config import resolve_log_level

        os.environ["LOG_LEVEL"] = "VERBOSE"
        with self.assertRaises(ValueError):
            resolve_log_level()

    def test_create_server_configures_logging_to_stderr(self) -> None:
        from word_mcp.server import create_server

        os.environ["LOG_LEVEL"] = "INFO"
        stdout_buf = io.StringIO()
        stderr_buf = io.StringIO()
        with redirect_stdout(stdout_buf), redirect_stderr(stderr_buf):
            create_server()

        self.assertEqual(stdout_buf.getvalue(), "")
        self.assertIn("Starting Word MCP Server", stderr_buf.getvalue())

    def test_exception_tool_logs_error_when_dev_mode_enabled(self) -> None:
        from word_mcp.exception.exception_tools import ExceptionTool
        from word_mcp.utils.logging_config import configure_logging

        os.environ["DEV_MODE"] = "true"
        stderr_buf = io.StringIO()
        with redirect_stderr(stderr_buf):
            configure_logging()
            res = ExceptionTool.handle_error(
                ValueError("bad input"),
                operation="unit_test",
            )

        self.assertEqual(res["status"], "error")
        self.assertIn("bad input", stderr_buf.getvalue())
        self.assertIn("unit_test", stderr_buf.getvalue())


if __name__ == "__main__":
    unittest.main()
