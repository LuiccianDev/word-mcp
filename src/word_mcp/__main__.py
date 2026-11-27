"""
Entry point for running word_mcp as a module.

This module provides the command-line interface for the Word MCP server
over stdio transport.
"""

from __future__ import annotations

import argparse
import sys

from word_mcp import __version__


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
    return parser.parse_args(args)


def main() -> None:
    """Main CLI entry point."""
    parse_args(sys.argv[1:])

    from word_mcp.server import create_server

    mcp = create_server()
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
