<div align="center">
  <h1>Word MCP Server</h1>
  <p>
    <em>High-performance Model Context Protocol (MCP) server for programmatic Word (.docx) document generation and manipulation.</em>
  </p>
</div>

<p align="center">
  <a href="https://pypi.org/project/luiccian-word-mcp/"><img src="https://img.shields.io/pypi/v/luiccian-word-mcp.svg" alt="PyPI Version" /></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.11+-3776AB.svg" alt="Python 3.11+" /></a>
  <a href="https://github.com/astral-sh/uv"><img src="https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json" alt="uv" /></a>
  <a href="https://github.com/LuiccianDev/word-mcp/actions/workflows/ci.yml"><img src="https://github.com/LuiccianDev/word-mcp/actions/workflows/ci.yml/badge.svg" alt="CI Build" /></a>
  <a href="https://github.com/LuiccianDev/word-mcp/actions/workflows/codeql.yml"><img src="https://github.com/LuiccianDev/word-mcp/actions/workflows/codeql.yml/badge.svg" alt="CodeQL" /></a>
  <a href="https://github.com/LuiccianDev/word-mcp/blob/master/LICENSE"><img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License" /></a>
  <a href="https://modelcontextprotocol.io"><img src="https://img.shields.io/badge/MCP-Compatible-brightgreen" alt="MCP Compatible" /></a>
</p>

**Word MCP Server** (`word-mcp`, PyPI: [`luiccian-word-mcp`](https://pypi.org/project/luiccian-word-mcp/)) connects AI coding agents (Claude Desktop, Claude Code, Cursor, VS Code / GitHub Copilot, Antigravity, Windsurf, Cline) to deterministic `.docx` generation, inspection, and protection workflows. By default, it exposes **9 tools in `compound` mode** (3 single-pass builders and 6 domain dispatchers covering 40 granular actions) to cut schema overhead and build full documents in one call.

---

## Quick Start

### 1. Run Directly with `uvx` (Recommended)

Requires Astral [`uv`](https://docs.astral.sh/uv/). No repository clone or manual virtual environment is needed:

```bash
# Run over local stdio transport for AI IDEs and desktop clients
uvx luiccian-word-mcp
```

### 2. Configure Your MCP Client (Claude Desktop)

Add the server to `claude_desktop_config.json` and point `WORD_MCP_DIRS` to the folder(s) the assistant may access:

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

> **Note:** On Windows, use forward slashes or escaped backslashes (for example, `"C:/Users/YourName/Documents"`). Separate multiple allowed directories with commas.

---

## Architecture & Tool Modes

### Selective Tool Loading (`WORD_MCP_TOOLS`)

| Mode | Tools Exposed | Behavior |
| :--- | :--- | :--- |
| `compound` (default) | 9 tools | Exposes 3 single-pass builders plus 6 domain routers (`word_document` through `word_property`) with `action="help"` schema discovery. |
| `full` / `granular` | 40 tools | Registers each atomic tool directly without a domain wrapper. |

### Default Compound Catalog (`WORD_MCP_TOOLS=compound`)

| Tool | Role | Actions / Capabilities |
| :--- | :--- | :--- |
| `build_document` | Single-Pass Builder | Create complete documents with presets, margins, page sizes, headers/footers, `PAGE` numbers, TOC, callouts, tables, images, and lists in 1 call. |
| `create_from_markdown` | Single-Pass Builder | Convert Markdown (headings, tables, nested lists, inline links, and GitHub alerts `> [!NOTE]`) into a styled `.docx` in 1 call. |
| `batch_add_content` | Single-Pass Appender | Append multiple structured content blocks atomically to an existing `.docx`. |
| `word_document` | Domain Dispatcher | `create`, `copy`, `info`, `text`, `outline`, `list`, `merge`, `paragraph_text`, `find_text`, `to_pdf`, `read_range`, `get_tables`, `help` |
| `word_content` | Domain Dispatcher | `add_heading`, `add_paragraph`, `add_picture`, `add_table`, `add_page_break`, `delete_paragraph`, `search_and_replace`, `add_toc`, `add_hyperlink`, `add_bookmark`, `batch_edit`, `help` |
| `word_format` | Domain Dispatcher | `format_text`, `create_style`, `format_table`, `add_header`, `add_footer`, `help` |
| `word_protection` | Domain Dispatcher | `protect`, `unprotect`, `restrict_editing`, `add_signature`, `verify`, `help` |
| `word_footnote` | Domain Dispatcher | `add_footnote`, `add_endnote`, `convert_to_endnotes`, `customize_style`, `help` |
| `word_property` | Domain Dispatcher | `get_properties`, `set_properties`, `set_page_layout`, `help` |

### Token Optimization: Single-Pass vs. Legacy

| Metric | Legacy Micro-Calls | Word MCP Single-Pass |
| :--- | :--- | :--- |
| **Tool Calls** | 15 - 35 calls per document | **1 single call** (`build_document`) |
| **Token Usage** | ~15,000 - 30,000 tokens | **~800 - 1,500 tokens** (90%+ savings) |
| **Disk Operations** | Multiple file writes | Single in-memory assembly and atomic `os.replace` |

Outline inspection (`get_document_outline`) extracts paragraph previews and table counts without dumping full body text into context.

---

## Visual Design Presets & Configuration

### Built-in Presets (`preset`)

| Preset | Font Family | Body Size | Line Spacing | Page Size | Margins | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `modern` *(Default)* | Calibri | 11.0 pt | 1.15 | Letter | 2.5 cm | Executive summaries, proposals, corporate reports |
| `formal` | Times New Roman | 12.0 pt | 1.50 | Letter | Top/Bottom/Right 2.5 cm, Left 3.0 cm | Legal contracts, official letters, meeting minutes |
| `report` | Calibri | 11.0 pt | 1.15 | A4 | 2.5 cm | Engineering deliverables, technical audits, structured reports |
| `academic` | Times New Roman | 12.0 pt | 2.00 (Double) | Letter | 2.54 cm (1 inch) | University papers, research articles, theses |
| `apa` | Times New Roman | 12.0 pt | 2.00 (Double) | Letter | 2.54 cm (1 inch) | APA 7th edition manuscripts with hanging-indent references |
| `compact` | Arial | 10.0 pt | 1.00 (Single) | Letter | 2.0 cm | Invoices, datasheets, internal memos |

### Environment Variables

| Variable | Default | Description |
| :--- | :--- | :--- |
| `WORD_MCP_DIRS` | `./documents` | Comma-separated allowed directories for reading, writing, listing, images, and PDF export across all path arguments (legacy fallback: `MCP_ALLOWED_DIRECTORIES`). |
| `WORD_MCP_TOOLS` | `compound` | Tool surface mode: `compound` (9 tools) or `full` / `granular` (40 tools) (legacy fallback: `WORD_MCP_TOOL_MODE`). |
| `WORD_MCP_LOG_LEVEL` | `INFO` | Logging verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`) written strictly to `stderr`. |
| `WORD_MCP_DEV_MODE` | `false` | Set `true` to default logging to `DEBUG` and emit full tracebacks to `stderr`. |

---

## Other MCP Clients & Local Development

<details>
<summary>Cursor, VS Code, Antigravity, Windsurf, Cline, Claude Code CLI, and Local `uv` Setup</summary>

### Cursor (`.cursor/mcp.json`)

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

### VS Code / GitHub Copilot (`.vscode/mcp.json`)

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

### Antigravity / Windsurf (`mcp_config.json`) & Cline (`cline_mcp_settings.json`)

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

### Claude Code CLI

```bash
claude mcp add word-mcp \
  -e WORD_MCP_DIRS="/path/to/your/documents" \
  -- uvx luiccian-word-mcp
```

### Local Development Checkout (`uv`)

```bash
git clone https://github.com/LuiccianDev/word-mcp.git
cd word-mcp
uv sync --all-groups
uv run ruff check src tests scripts
uv run ruff format --check src tests scripts
uv run mypy src
uv run pytest
```

To point an MCP client at a local checkout instead of PyPI:

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/word-mcp", "luiccian-word-mcp"],
      "env": {
        "WORD_MCP_DIRS": "/path/to/your/documents"
      }
    }
  }
}
```

</details>

---

## Documentation, Skills & Governance

| Resource | Purpose |
| :--- | :--- |
| [Tools Reference](https://github.com/LuiccianDev/word-mcp/blob/master/docs/TOOLS.md) | Complete signatures, return shapes, and parameter contracts (`docs/TOOLS.md`) |
| [MCPB Deployment Guide](https://github.com/LuiccianDev/word-mcp/blob/master/docs/MCPB.md) | Building and installing desktop `.mcpb` bundles |
| [Agent Skills (`skills/`)](https://github.com/LuiccianDev/word-mcp/tree/master/skills) | Modular LLM skills: `word-document-creator`, `word-document-editor`, `word-document-security` |
| [Security Policy](https://github.com/LuiccianDev/word-mcp/blob/master/SECURITY.md) | Sandbox isolation, `.protection` sidecar security, and vulnerability disclosure |
| [Contributing Guide](https://github.com/LuiccianDev/word-mcp/blob/master/CONTRIBUTING.md) | Quality gate, commit conventions, and automated release flow |
| [AI Assistant Rules](https://github.com/LuiccianDev/word-mcp/blob/master/AGENTS.md) | Repository architecture rules, boundary contracts, and token optimization policy |
| [Changelog](https://github.com/LuiccianDev/word-mcp/blob/master/CHANGELOG.md) | Version history following Keep a Changelog |

Licensed under the [MIT License](https://github.com/LuiccianDev/word-mcp/blob/master/LICENSE). Maintained by [LuiccianDev](https://github.com/LuiccianDev).
