# AGENTS.md — Rules for AI Assistants

Project: **word-mcp** (High-Performance Word .docx MCP Server • PyPI: `luiccian-word-mcp`).

---

## 0. Language, Aesthetics & Documentation Policy

- **Repository Language:** All official repository documentation, code comments, pull requests, issue templates, workflows, and git commits **MUST** be written in **Technical English**.
- **User Interaction:** When interacting with the user, the assistant should reply in **Spanish** (or the user's preferred language) while preserving technical names in English and maintaining a formal engineering tone.
- **Strict Aesthetic Standard (Zero Informal Emojis):** Prohibit informal emojis (such as rockets, party poppers, fire, tool icons) in markdown documentation, section headers, release notes, issue templates, and commit messages. Maintain a sober engineering format utilizing markdown tables and GitHub alerts (`> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`).

---

## 1. Zero-Install Tier 1 Priority

When assisting users or configuring clients (Claude Desktop, Cursor, Cline, Copilot), assistants must adhere to the following priority tiers:

- **Tier 1 (Recommended — Zero-Install):** Prioritize direct execution with `uvx luiccian-word-mcp`. Users do not need to clone the repository or manually manage Python virtual environments.
  ```json
  {
    "mcpServers": {
      "word-mcp": {
        "command": "uvx",
        "args": ["luiccian-word-mcp"],
        "env": {
          "WORD_MCP_DIRS": "/path/to/documents"
        }
      }
    }
  }
  ```
- **Tier 2 (Advanced / Local Development):** For contributors developing features or modifying tools locally, use Astral `uv`:
  ```bash
  uv sync --all-groups
  uv run luiccian-word-mcp
  ```
- **Runtime Environment Variables:**
  - `WORD_MCP_DIRS`: Comma-separated allowed directories (legacy fallback: `MCP_ALLOWED_DIRECTORIES`).
  - `WORD_MCP_TOOLS`: Tool surface mode, defaulting to `compound` (9 tools); set `full` (or `granular`) for all 40 tools (legacy fallback: `WORD_MCP_TOOL_MODE`).
  - `WORD_MCP_LOG_LEVEL`: Logging verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`) written strictly to `stderr`.
  - `WORD_MCP_DEV_MODE`: Set `true` to default logging to `DEBUG` and emit full exception tracebacks to `stderr`.

---

## 2. Single-Pass Token Optimization & Compound Dispatch Directives

- **Prohibition of Chatty Micro-Iterations:** AI assistants using `word-mcp` must avoid making consecutive micro-calls (e.g., calling `create_document`, then 15 separate calls to `add_heading`, `add_paragraph`, `add_table`, `set_page_numbering`, and `save_document`). This legacy anti-pattern consumes tens of thousands of tokens and multiple conversational turns.
- **Mandatory Single-Pass Usage:** Always prefer:
  1. `build_document`: Construct complete, publication-ready documents with headers, footers, page numbers (`PAGE`), margins, typography presets (`modern`, `formal`, `report`, `academic`, `apa`, `compact`), headings, markdown styling, callout boxes, tables, and lists in **one single MCP tool call** (>90% token savings).
  2. `create_from_markdown`: Convert structured markdown text directly into a Word document in one call.
  3. `batch_add_content`: Batch-append sections, tables, callouts, and lists to existing documents in a single atomic pass.
  4. `word_content(action="batch_edit", ...)`: Apply multiple text/paragraph replacements, insertions, and deletions atomically in one pass.
- **Lightweight Outline & Bounded Inspection:** When reading or modifying large documents, use `word_document(action="outline", ...)` (`get_document_outline`), `word_document(action="read_range", ...)` (`read_document_range`), or `word_document(action="get_tables", ...)` (`get_document_tables`) first to inspect structure, section slices, and paginated table rows without dumping full body text into the context window.
- **Single-Round-Trip Schema Discovery:** In default `compound` mode (9 tools), call `action="help"` with no `params` on any domain tool (`word_document`, `word_content`, `word_format`, `word_protection`, `word_footnote`, `word_property`) to retrieve the full `required` and `optional` parameter schema for every action in that domain in one call.

---

## 3. Package Scope & Isolation Boundary

The distribution boundary between the PyPI package and the repository is strictly defined:

- **PyPI Package (`luiccian-word-mcp`):** Bundles **exclusively** the runtime code located in [`src/word_mcp/`](file://src/word_mcp/), accompanied by `README.md`, `LICENSE`, and `pyproject.toml`.
- **Repository Scope:** Test suites (`tests/`), automation scripts (`scripts/`), agent skills (`skills/`), documentation manuals (`docs/`), and GitHub Actions workflows (`.github/`) belong solely to repository infrastructure and are explicitly excluded from the PyPI wheel and sdist distributions via Hatchling build targets.

---

## 4. Official Naming & Hierarchy Standards

To guarantee consistency, cleanliness, and maintainability across the project:

### 1. Root & Module Hierarchy
- **Source Module:** Python source resides strictly in `src/word_mcp/`.
- **Package Name on PyPI:** `luiccian-word-mcp`.
- **Command-line Script:** `luiccian-word-mcp` is registered in `pyproject.toml` and points to `word_mcp.__main__:main`.
- **Support Folders (Lowercase):**
  - `docs/` (technical documentation and tool specifications)
  - `scripts/` (release, versioning, and build automation)
  - `skills/` (modular skills for LLMs: `word-document-creator`, `word-document-editor`, `word-document-security`)
  - `tests/` (automated test suites)
  - `.github/` (CI/CD and security workflows)

### 2. Python Code Standards (PEP 8)
- **Module & File Names:** Strictly `snake_case.py` (e.g., `batch.py`, `document_tools.py`, `response_models.py`). Never use PascalCase or camelCase for `.py` files.
- **Classes:** `PascalCase` (`DocumentContext`, `HeadingResult`).
- **Functions & Variables:** `snake_case` (`build_document()`, `style_preset`).
- **Constants:** `SCREAMING_SNAKE_CASE` (`DEFAULT_PAGE_SIZE`, `MAX_TABLE_ROWS`).
- **Automation Scripts:** Strictly `kebab-case.py` in `scripts/` (e.g., `bump-version.py`, `extract-release-notes.py`).

### 3. Boundary Validation, Atomic Persistence & Error Contract (v3.2.0+)
- **Sandbox-First Boundary (`@tool_boundary` & `PathRule`):** All tool path parameters (primary, optional, `list[str]`, directory, and secondary arguments such as `destination_filename`, `source_filenames`, `output_filename`, `image_path`, and batch `image` elements) must be validated through `word_mcp.validation.document_validators` (`validate_path` / `resolve_sandboxed_path`) **before** any filesystem existence probe, directory creation, or ZIP check.
- **Single-Open & Atomic Writes (`atomic_write_target`):** Pre-flight checks use `_fast_check_ooxml_zip` (`word/document.xml` and `[Content_Types].xml` inspection without XML DOM parsing) so `document_context` opens `docx.Document` only once per tool call. All `.docx` and `.protection` writes must stage to a sibling temporary file in the target directory and replace atomically via `atomic_write_target` (`os.replace`).
- **Canonical Error Envelope (`OperationError`):** Every tool and validator failure (outside FastMCP `ToolError` dispatch diagnostics) must be routed through `ExceptionTool.handle_error` and return the `OperationError` schema (`status="error"`, `error_type`, `message`, `suggestion`, `recoverable`). Never return ad-hoc `{"error", "hint"}` or bare `{"status": "error", "message": ...}` dicts.
- **Compound Dispatch (`ActionSpec`):** Tools dispatched by `compound_tools.py` must declare required parameters without placeholder `""` defaults so `ActionSpec` and `docs/TOOLS.md` stay synchronized via `scripts/sync-tools-reference.py`. Calling `action="help"` with no `params` returns the full required/optional parameter schema for all actions in the domain group in a single call.

---

## 5. Optimization of Minutes and Compute in GitHub Actions

To preserve compute quotas and prevent unnecessary runner execution in GitHub Actions:

1. **Exclusion of Non-executable Files (`paths-ignore`):**
   - Workflows (`ci.yml`, `codeql.yml`) must ignore changes that exclusively touch documentation (`**.md`, `docs/**`, `skills/**`), licensing, or git metadata (`.gitignore`, `.gitattributes`).
2. **Mandatory Use of `[skip ci]`:**
   - In commits that are strictly documentary, markdown-related, or metadata refinements that do not affect runtime Python code or tests, always include `[skip ci]` or `[ci skip]` in the commit message.
3. **Automatic Cancellation of Stale Runs (`concurrency`):**
   - Workflows must maintain `concurrency: cancel-in-progress: true` to immediately terminate superseded jobs when new commits are pushed to the same branch.

---

## 6. Semantic Versioning & CHANGELOG Maintenance

1. **SemVer 2.0.0 & Keep a Changelog:**
   - Adhere strictly to [Semantic Versioning](https://semver.org/) (`MAJOR.MINOR.PATCH`) and [Keep a Changelog](https://keepachangelog.com/).
   - Valid changelog subsections: `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, `Security`.
2. **Unreleased Section (`[Unreleased]`):**
   - Active developments not yet published with a git tag must be recorded under `## [Unreleased]`.
3. **Atomic 1-Click Releases:**
   - Do not perform manual version tagging or ad-hoc PyPI publishing. Use the automated workflow `.github/workflows/bump-and-release.yml` with `scripts/bump-version.py` to atomically bump version in `pyproject.toml`, promote changelog notes, create the git tag `v*`, and deploy to PyPI via Trusted Publishing OIDC.

---

## 7. Graphify Knowledge Graph & Token-Saving Protocol

- **Location:** `graphify-out/` (`graph.json`, `GRAPH_REPORT.md`, `graph.html`) and local plugin config in `.agents/plugins/graphify/`.
- **When to Use:** Query Graphify MCP tools (`query_graph`, `get_node`, `get_neighbors`, `shortest_path`, `god_nodes`) or inspect `graphify-out/GRAPH_REPORT.md` before performing broad recursive searches.
- **Maintenance:** Update deterministically with zero API token cost after modifying code in `src/word_mcp/` or `docs/`:
  ```powershell
  uvx --from graphifyy graphify update .
  ```
