"""
Main entry point for the Word Document MCP Server.

Acts as the central controller for the MCP server that handles Word document operations.
"""

from fastmcp import FastMCP

from word_mcp import __version__
from word_mcp.tools.register_tools import get_tool_mode, register_all_tools
from word_mcp.utils.logging_config import configure_logging, is_dev_mode
from word_mcp.utils.security import get_allowed_directories


def create_server() -> FastMCP:
    """Create and configure the Word MCP Server."""
    logger = configure_logging()
    tool_mode = get_tool_mode()
    allowed_dirs = get_allowed_directories()
    dirs_list = "\n".join([f"- `{d}`" for d in allowed_dirs])

    if tool_mode == "compound":
        tool_naming_section = (
            "Mode: `compound` (9 tools). Prefer single-pass `build_document`, "
            "`create_from_markdown`, and `batch_add_content` for generation; "
            "for large documents inspect with `word_document` actions (`outline`, "
            "`read_range`, `get_tables`) and batch edits with `word_content` "
            "(`batch_edit`). Pass args in `params`; call `action='help'` (no `params`) "
            "to inspect all action schemas in one call."
        )
    else:
        tool_naming_section = (
            "Mode: `full` (40 unprefixed tools, registered once each). "
            "Prefer `build_document`, `create_from_markdown`, and `batch_add_content`; "
            "for large documents inspect with `get_document_outline` and "
            "`find_text_in_document` before calling `get_document_text`."
        )

    mcp = FastMCP(
        name="word_mcp",
        instructions=(
            "Word MCP Server (.docx manipulation over local stdio).\n"
            f"- {tool_naming_section}\n"
            "- Security: strictly restricted to these allowed directories "
            "(default here when no absolute path is given):\n"
            f"{dirs_list}\n"
            "- Protection uses `msoffcrypto` encryption; signature markers store a "
            "SHA-256 integrity hash (not a cryptographic signature). "
            "Footnote-to-endnote conversion is one-way."
        ),
        on_duplicate="error",
    )

    register_all_tools(mcp)
    logger.info(
        "Starting Word MCP Server v%s (dev_mode=%s)", __version__, is_dev_mode()
    )
    logger.debug("Allowed directories: %s", allowed_dirs)

    return mcp


if __name__ == "__main__":
    mcp = create_server()
    mcp.run(transport="stdio")
