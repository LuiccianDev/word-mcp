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
            "Active mode (`WORD_MCP_TOOLS=compound`, default): 9 tools are registered. "
            "Prefer single-pass `build_document`, `create_from_markdown`, and `batch_add_content` "
            "for document generation. For granular operations, use the 6 domain dispatchers "
            "(`word_document`, `word_content`, `word_format`, `word_protection`, `word_footnote`, "
            "`word_property`) and call `action='help'` with no `params` to discover every "
            "action's required and optional schema in one round trip."
        )
    else:
        tool_naming_section = (
            "Active mode (`WORD_MCP_TOOLS=full`): 40 unprefixed tools are registered "
            "(for example, `create_document`, `build_document`, `create_from_markdown`, "
            "`batch_add_content`). Each tool is registered exactly once."
        )

    mcp = FastMCP(
        name="word_mcp",
        instructions=f"""# Word MCP Server

Word MCP Server provides comprehensive Word document (.docx) manipulation capabilities for AI assistants.

## Capabilities

### Document Operations
- **Create documents**: Generate new Word documents with optional metadata (title, author)
- **Read documents**: Extract text, information, and structure from existing documents
- **Copy documents**: Duplicate documents with automatic naming
- **Merge documents**: Combine multiple documents into one
- **Convert to PDF**: Export documents to PDF format

### Content Management
- **Add content**: Headings, paragraphs, tables, pictures, page breaks
- **Search & replace**: Find and replace text throughout documents
- **Delete content**: Remove paragraphs by index

### Formatting
- **Text formatting**: Bold, italic, underline, colors, font size, font family
- **Custom styles**: Create reusable character and paragraph styles
- **Table formatting**: Borders, shading, header rows

### Footnotes & Endnotes
- **Add footnotes**: Reference notes at page bottom
- **Add endnotes**: Reference notes at document end
- **Convert footnotes to endnotes**: Transform existing footnotes into endnotes (one-way)
- **Customize styling**: Configure footnote/endnote appearance

### Document Protection
- **Password protection**: Encrypt documents with passwords
- **Restricted editing**: Limit editing to specific sections
- **Signature markers**: Add a visible signature block and a SHA-256 content hash
- **Verification**: Check protection metadata and signature-marker status

## Usage Patterns

### Automation
- Generate reports from templates
- Batch process multiple documents
- Extract data from existing documents
- Apply consistent formatting across documents

### Document Creation
- Create new documents with structured content
- Add tables with data
- Insert images
- Build tables of contents

## Security

- **ALLOWED DIRECTORIES:**
  You are STRICTLY RESTRICTED to reading, creating, or modifying files ONLY within the following permitted directories:
{dirs_list}

  If the user asks to create or manipulate a document without specifying an absolute path, you MUST default to one of the above allowed directories.
- Document protection encrypts the OOXML container with `msoffcrypto`
- Signature markers detect text changes; they are not cryptographic signatures

## Tool Naming

{tool_naming_section}""",
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
