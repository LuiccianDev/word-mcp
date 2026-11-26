"""
Entry point for running word_mcp as a module.

This module provides the command-line interface for the Word MCP server,
supporting both stdio and SSE transport modes.
"""

from __future__ import annotations

import argparse
import sys

from word_mcp import __version__


_LOOPBACK_HOSTS: frozenset[str] = frozenset({"127.0.0.1", "localhost", "::1"})


def parse_args(args: list[str] | None = None) -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        prog="luiccian-word-mcp",
        description="Word MCP Server - High-performance Word document manipulation via MCP",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="Transport protocol to use (default: stdio)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host address for SSE transport (default: 127.0.0.1)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port number for SSE transport (default: 8000)",
    )
    parser.add_argument(
        "--allow-remote",
        action="store_true",
        default=False,
        help="Allow binding SSE transport to non-loopback interfaces (use only behind an authenticated reverse proxy)",
    )
    return parser.parse_args(args)


def main() -> None:
    """Main CLI entry point."""
    parsed = parse_args(sys.argv[1:])
    if (
        parsed.transport == "sse"
        and parsed.host.strip().lower() not in _LOOPBACK_HOSTS
        and not parsed.allow_remote
    ):
        raise SystemExit(
            "Error: Binding SSE transport to a non-loopback host requires --allow-remote "
            "because SSE transport has no built-in authentication."
        )

    from word_mcp.server import create_server

    mcp = create_server()

    if parsed.transport == "sse":
        mcp.run(transport="sse", host=parsed.host, port=parsed.port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
