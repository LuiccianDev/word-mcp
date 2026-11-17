<div align="center">
  <h1>Word MCP Server</h1>
  <p>
    <em>High-performance Model Context Protocol (MCP) server for programmatic Word (.docx) document generation and manipulation</em>
  </p>

  [![PyPI Version](https://img.shields.io/pypi/v/luiccian-word-mcp.svg)](https://pypi.org/project/luiccian-word-mcp/)
  [![Python Version](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/)
  [![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)
  [![CodeQL](https://github.com/LuiccianDev/word-mcp/actions/workflows/codeql.yml/badge.svg)](https://github.com/LuiccianDev/word-mcp/actions/workflows/codeql.yml)
  [![Security Policy](https://img.shields.io/badge/security-policy-blue.svg)](SECURITY.md)
  [![Rules for AI](https://img.shields.io/badge/rules-AGENTS.md-blue.svg)](AGENTS.md)
  [![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)

</div>

**Word MCP Server** (`word-mcp`, PyPI: [`luiccian-word-mcp`](https://pypi.org/project/luiccian-word-mcp/)) is a high-performance Python server implementing the Model Context Protocol (MCP) to provide token-optimized Word document (`.docx`) creation, batch synthesis, and inspection capabilities for AI assistants (Claude Desktop, Claude Code, Cursor, VS Code / GitHub Copilot, Antigravity, Windsurf, Cline).

Designed from the ground up for maximum token efficiency, it replaces chatty micro-iterations with single-pass document construction, saving over 90% of model tokens and conversational turns.

---

## Key Architectural Highlights

- **Compound Mode by Default (v3.0.0)**: Registers 9 tools by default (`build_document`, `create_from_markdown`, `batch_add_content`, plus 6 action-dispatched domain tools: `word_document`, `word_content`, `word_format`, `word_protection`, `word_footnote`, `word_property`) to reduce tool-schema context overhead by 77.5%. Set `WORD_MCP_TOOLS=full` (or `granular`) to expose all 40 individual tools.
- **Single-Pass Document Builder**: Create complete, publication-ready documents (headers, footers, dynamic `PAGE` numbers, margins, headings, markdown styling, styled callout boxes, tables, TOCs, images, and lists) in **one single MCP call** (`build_document`).
- **Markdown Direct Generation**: Turn structured Markdown into styled Word documents in one pass (`create_from_markdown`), including tables, nested lists, inline links, and GitHub-style alert callouts (`> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`, `> [!CAUTION]`).
- **Visual Design Presets**: Built-in typography and spacing standards (`modern`, `formal`, `academic`, `compact`, `report`, `apa`) with optional `font_name`, `font_size`, and `page_size` overrides.
- **Lightweight Outline Inspection**: Extract paragraph previews and table counts without dumping full body text (`get_document_outline`), saving thousands of context tokens.
- **Batch Content Ingestion**: Append multiple sections, tables, callouts, and lists to existing documents in a single atomic operation (`batch_add_content`).
- **Protection & Security**: Password encryption, editing restrictions, file integrity checks, and headless PDF export.
- **Complete Tool Reference**: See [`docs/TOOLS.md`](docs/TOOLS.md) for every tool signature, return shape, and parameter contract.
- **Zero-Install Execution**: Instant launch with `uvx luiccian-word-mcp`.

---

## Quick Start with `uvx` (Recommended)

No manual clone or virtual environment management required. Run the server directly using Astral [`uv`](https://docs.astral.sh/uv/):

```bash
# Run with stdio transport (default for AI IDEs)
uvx luiccian-word-mcp

# Or run with SSE transport
uvx luiccian-word-mcp --transport sse --host 127.0.0.1 --port 8000
```

---

## MCP Client Configuration

Configure your AI client by passing environment variables directly in the client's `env` block (no `.env` file is required). By default, the server starts in `compound` mode (9 tools); set `"WORD_MCP_TOOLS": "full"` in `env` to expose all 40 granular tools.

### 1. Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "WORD_MCP_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

> [!TIP]
> **Windows path example:** Use forward slashes or escaped backslashes (e.g., `"C:/Users/YourName/Documents"`). You can allow multiple roots separated by commas.
>
> **Local checkout:** If running from a local git repository instead of PyPI:
> ```json
> {
>   "mcpServers": {
>     "word-mcp": {
>       "command": "uv",
>       "args": ["run", "--directory", "/path/to/word-mcp", "luiccian-word-mcp"],
>       "env": {
>         "WORD_MCP_DIRS": "/path/to/your/documents"
>       }
>     }
>   }
> }
> ```

### 2. Cursor (`.cursor/mcp.json`)

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "WORD_MCP_DIRS": "${workspaceFolder}"
      }
    }
  }
}
```

### 3. VS Code / GitHub Copilot (`.vscode/mcp.json`)

VS Code uses a top-level `"servers"` key and explicit `"type": "stdio"`:

```json
{
  "servers": {
    "word-mcp": {
      "type": "stdio",
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "WORD_MCP_DIRS": "${workspaceFolder}"
      }
    }
  }
}
```

### 4. Antigravity / Windsurf (`mcp_config.json`)

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "WORD_MCP_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

### 5. Cline / Roo Code (`cline_mcp_settings.json`)

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "WORD_MCP_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

### 6. Claude Code CLI

```bash
claude mcp add word-mcp \
  -e WORD_MCP_DIRS="/path/to/your/documents" \
  -- uvx luiccian-word-mcp
```

---

## Environment Variables

All runtime configuration is supplied via standard OS environment variables or the MCP client's `env` object:

| Variable | Default | Description |
| :--- | :--- | :--- |
| `WORD_MCP_DIRS` | *(current directory)* | Comma-separated list of allowed root directories for reading, writing, and listing `.docx` files. Legacy `MCP_ALLOWED_DIRECTORIES` remains supported as a fallback. |
| `WORD_MCP_TOOLS` | `compound` | Tool surface mode: `compound` (9 tools: 3 single-pass + 6 domain dispatchers) or `full` (or `granular`, all 40 tools). Legacy `WORD_MCP_TOOL_MODE` remains supported as a fallback. |
| `LOG_LEVEL` | `INFO` | Logging verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`) written strictly to `stderr`. Also accepts `WORD_MCP_LOG_LEVEL`. |
| `DEV_MODE` | `false` | Set to `true` to default logging to `DEBUG` and log full exception tracebacks to `stderr`. Also accepts `WORD_MCP_DEV_MODE`. |

---

## Token Optimization: Single-Pass vs. Legacy

| Metric | Legacy Micro-Calls | Word MCP Single-Pass |
| :--- | :--- | :--- |
| **Tool Calls** | 15 - 35 calls per document | **1 single call** (`build_document`) |
| **Token Usage** | ~15,000 - 30,000 tokens | **~800 - 1,500 tokens** (90%+ savings) |
| **Disk Operations** | Multiple file writes | Single in-memory assembly and write |
| **Generation Time** | 30 - 60 seconds | Single in-memory assembly |

---

## Visual Design Presets

| Preset | Font Family | Body Size | Line Spacing | Margins | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `modern` *(Default)* | Calibri | 11.0 pt | 1.15 | 2.5 cm | Executive summaries, business proposals, corporate reports |
| `formal` | Times New Roman | 12.0 pt | 1.50 | Top/Bottom 2.5 cm, Left 3.0 cm | Legal contracts, official letters, meeting minutes |
| `academic` | Times New Roman | 12.0 pt | 2.00 (Double) | 2.54 cm (1 inch) | University papers, research articles, theses (APA/MLA) |
| `compact` | Arial | 10.0 pt | 1.00 (Single) | 2.0 cm | Invoices, datasheets, internal memos |

---

## Modular Skills Ecosystem

The repository provides 3 modular, specialized skills in `skills/` for LLMs:

- **`word-document-creator`**: Instructs the model to build complete documents in 1 single pass with headers, footers, pagination, tables, callouts, and presets.
- **`word-document-editor`**: Focuses on lightweight outline inspection, batch mutations, search and replace, and document merging.
- **`word-document-security`**: Handles password locking, editing restrictions, and PDF export for distribution.

---

## Security

- **Security Policy**: See [`SECURITY.md`](SECURITY.md) for vulnerability disclosure, protection metadata, and threat modeling.
- **Rules for AI Assistants**: See [`AGENTS.md`](AGENTS.md) for official development directives, coding standards, and token optimization rules.
- **Directory Isolation**: Set `WORD_MCP_DIRS` to restrict document operations strictly to designated folders. Path canonicalization prevents directory traversal.
- **Trusted Publishing**: Tag-driven PyPI releases (`v*.*.*`) use GitHub Actions OIDC and PEP 740 artifact attestation.

---

## Development

```bash
# Clone the repository
git clone https://github.com/LuiccianDev/word-mcp.git
cd word-mcp

# Install dependencies with Astral uv
uv sync --all-groups

# Run server locally
uv run luiccian-word-mcp

# Run linting and type checking
uv run ruff check src tests scripts
uv run mypy src

# Run test suite
uv run pytest
```

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

<div align="center">
  <p><strong>Word MCP Server</strong> • Maintained by <a href="https://github.com/LuiccianDev">LuiccianDev</a></p>
</div>
