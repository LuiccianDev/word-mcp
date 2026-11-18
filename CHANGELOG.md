# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [3.2.0] - 2026-11-18

### Added
- Added `SecurityError` and contextual `suggestion`/`details` fields on domain exceptions (`ContextualFileNotFoundError`, `DocumentProcessingError`, `ValidationError`, `FileOperationError`, `StyleError`, `ConfigurationError`).
- Added declarative `PathAccess` / `PathRule` boundary validation (`tool_boundary`) and `atomic_write_target` same-directory temporary file replacement (`os.replace`) across `document_context`, document creation/copy/merge, and protection/unprotect writers.
- Added `resolve_sandboxed_path` with empty, whitespace, and NUL-byte validation and `ActionSpec` caching in `compound_tools` returning full per-action parameter schemas in a single `action="help"` call.

### Changed
- Unified tool, validator, and batch error payloads under the canonical `OperationError` schema (`status`, `error_type`, `message`, `suggestion`, `recoverable`).
- Replaced full `Document()` XML DOM pre-flight parsing in validators with fast OOXML ZIP inspection (`_fast_check_ooxml_zip`) and pure `core_extract_document_text` / `core_get_document_structure` helpers for single-open reads.
- Canonicalized `filename: str` as a required parameter on `add_hyperlink`, `add_bookmark`, `add_header`, `add_footer`, `set_page_layout`, `get_core_properties`, and `set_core_properties` (preserving `file_path` alias normalization in compound mode) and added boolean-as-numeric rejection for `int`/`float` parameters.
- Updated `verify_document` to tolerate encrypted `.docx` files and report verification status via `ProtectionResult`.

### Fixed
- Fixed `msoffcrypto.OfficeFile.is_encrypted()` method invocation and password verification in `core_protect_document`, `core_unprotect_document`, and `core_verify_document_protection`.
- Fixed Windows file-locking misattribution in `_fast_check_ooxml_zip` by opening `ZipFile` directly so `PermissionError` surfaces as `file_operation`.
- Fixed `batch_tools` and `compound_tools` alias/conflict handling and ensured every compound group description includes a valid `example_params` payload.

### Security
- Enforced `WORD_MCP_DIRS` sandbox containment across all read/write/create decorators and secondary path parameters (`destination_filename`, `source_filenames`, `output_filename`, `image_path`, `directory`, and `build_document`/`batch_add_content` image elements).

### Removed
- Removed unused `ExceptionTool.wrap_tool_call`, `validate_docx_file`, and `check_file_writeable` helpers.

---

## [3.1.0] - 2026-11-18

### Added
- Styled GitHub-style callout boxes (`note`, `tip`, `important`, `warning`, `caution`) in `build_document`, `batch_add_content`, and `create_from_markdown` (`> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`, `> [!CAUTION]`) rendered as shaded 1x1 tables with colored left accent borders.
- Centralized `stderr`-only logging configuration (`WORD_MCP_LOG_LEVEL`) and developer diagnostics mode (`WORD_MCP_DEV_MODE`).

### Changed
- Standardized default `WORD_MCP_TOOLS` mode name to `compound` (9 tools) and removed the `compact` tool-mode alias; `full` (or `granular`) remains available for all 40 tools.
- Streamlined `README.md` to match the concise client configuration and environment variable reference format.

### Removed
- Removed unused Docker artifacts (`Dockerfile`, `.dockerignore`, `docs/Docker.md`) and untracked `docs/superpowers/`.

---

## [3.0.0] - 2026-11-17

### Changed
- **BREAKING**: Default tool registration mode is now `compact` (9 tools: `build_document`, `create_from_markdown`, `batch_add_content`, plus the 6 domain tools `word_document`, `word_content`, `word_format`, `word_protection`, `word_footnote`, and `word_property`) instead of 40 granular tools. Set `WORD_MCP_TOOLS=full` (or `granular`) to restore all 40 individual tools.
- Unified and simplified configuration environment variables across code, Docker, MCPB manifest, and documentation: `WORD_MCP_DIRS` replaces `MCP_ALLOWED_DIRECTORIES`, and `WORD_MCP_TOOLS` (accepting `compact`/`compound` and `full`/`granular`) replaces `WORD_MCP_TOOL_MODE`. Both legacy variable names remain supported as fallbacks for backward compatibility.

---

## [2.1.0] - 2026-11-16

### Added
- Opt-in compound tool registration mode (`WORD_MCP_TOOL_MODE=compound`) exposing a 9-tool hybrid surface (`build_document`, `create_from_markdown`, `batch_add_content`, plus `word_document`, `word_content`, `word_format`, `word_protection`, `word_footnote`, and `word_property`) to reduce tool-schema context overhead for clients with strict tool-count limits.
- Built-in `action="help"` schema inspection, `difflib` close-match suggestions for unknown actions/parameters, required/optional parameter listings, and Pydantic-backed parameter type validation/coercion in compound tools.
- Actionable `hint` guidance in document validator error responses (`_validate_docx_path`, `_validate_docx_path_shallow`, `validate_file_write`, `validate_docx_write`, `check_file_writeable`).

### Fixed
- Reordered `ExceptionTool.EXCEPTION_MAPPING` so `FileNotFoundError` and `PermissionError` precede `IOError` and `OSError`, preventing inheritance shadowing.

---

## [2.0.2] - 2026-11-10

### Added
- Documentation contract tests for tool signatures, return shapes, skills, version parity, packaging, and repository hygiene.
- `scripts/sync-tools-reference.py` to regenerate `docs/TOOLS.md` signature blocks from the registered tools.
- `src/word_mcp/py.typed` and the `Typing :: Typed` classifier.
- `CONTRIBUTING.md`, GitHub issue and pull request templates, and `.gitattributes`.

### Changed
- Aligned `docs/TOOLS.md`, `README.md`, `SECURITY.md`, `docs/MCPB.md`, and the three skills with the implemented runtime behavior.
- `scripts/bump-version.py` now updates the `docs/TOOLS.md` version footer.
- CI and the VS Code workspace now share the same `src`, `tests`, and `scripts` scope.
- **BREAKING**: MCP tools are now registered with a single, unprefixed name (for example, `build_document` and `create_document`). The duplicated `word_*` aliases have been removed, reducing the registered tool count from 80 to 40.
- Docker image installs runtime dependencies from `uv.lock`, launches the `luiccian-word-mcp` entry point, and bundles LibreOffice Writer for headless PDF conversion.
- CI quality gates for Ruff and mypy are now blocking, and the `tests/` and `scripts/` directories are linted.
- `scripts/bump-version.py` synchronizes `manifest.json` and validates custom version arguments.
- Package metadata migrated to the SPDX license expression (PEP 639) and Python 3.14 added to the support matrix.

### Removed
- Deleted the unused `word_mcp.utils.extended_document_utils` module and the unused `get_document_properties`, `find_paragraph_by_text`, and `find_and_replace_text` helpers.
- **BREAKING**: Deleted the `extended_document_tools`, `header_footer_tools`, and `link_tools` compatibility shims. Their functions now live in `document_tools`, `format_tools`, and `content_tools` respectively.
- Removed the Pydantic v1 fallback helper `_dump` from `batch_tools`; response models now call `model_dump()` directly.

### Fixed
- `manifest.json` now launches the `luiccian-word-mcp` console entry point instead of `src/word_mcp/server.py`.
- Corrected stale module docstrings in the validators package and the `ensure_docx_extension` comment.
- Corrected the tool names documented in `docs/TOOLS.md` and the document editor skill reference.
- Replaced legacy "Office MCP" branding and `mcp_word_office` links with the `word-mcp` identity.
- Restored the VS Code tasks and debug configuration to the `word_mcp` module.
- Aligned `manifest.json` and the `__version__` fallback with the published 2.0.1 release.

### Security
- Resolved CodeQL `py/weak-sensitive-data-hashing` alerts by replacing plain SHA-256 password hashing with PBKDF2-HMAC-SHA256 (100,000 iterations), unique cryptographic salts, and constant-time comparison in protection and unprotect modules.

---

## [2.0.1] - 2026-11-09

### Fixed
- Published the canonical `luiccian-word-mcp` console entry point so `uvx luiccian-word-mcp` resolves correctly.
- Aligned the distribution metadata with the renamed `word_mcp` package.

---

## [2.0.0] - 2026-11-07

### Added
- **Single-Pass Batch Document Builder**: High-performance engine in `word_mcp.core.batch` capable of generating complete Word documents with headers, footers, page numbering (`PAGE`), margins, typography, headings, paragraphs with markdown formatting, and tables in **1 single MCP call** (`word_build_document`).
- **Markdown Direct Generation**: `word_create_from_markdown` for immediate styled Word document generation directly from raw markdown strings.
- **Batch Content Ingestion**: `word_batch_add_content` to append multi-element sections into existing documents without conversational round-trips.
- **Visual Design Presets**: Native layout presets (`modern`, `formal`, `academic`, `compact`) standardizing font families, sizes, line spacing (1.15, 1.5, 2.0), paragraph spacing, and margins.
- **Modular Skills Ecosystem**: 3 focused skills in `skills/`:
  - `word-document-creator`: Single-pass creation patterns and block element cheatsheets.
  - `word-document-editor`: Lightweight outline inspection, batch mutations, search & replace, and document merging.
  - `word-document-security`: Document protection, passwords, editing restrictions, and PDF conversion.
- **Astral `uv` Automated Workflows**:
  - `ci.yml`: Fast multi-version testing and linting with cache and `paths-ignore` (skipping documentation edits to save runner minutes).
  - `release.yml`: Automated GitHub release generation with binary artifacts (.whl, .tar.gz) and automatic changelog note extraction.
  - `publish-pypi.yml`: Trusted Publishing OIDC deployment to PyPI with PEP 740 cryptographic attestation.
- **Automation Scripts**: `scripts/extract-release-notes.py` and `scripts/bump-version.py`.

### Changed
- **Package Renaming**: Renamed core package to `word_mcp` and project to `word-mcp`.
- **Dynamic Version Resolution**: `word_mcp.__version__` resolves dynamically via `importlib.metadata` with clean fallback.
- **Tool Architecture Consolidation**: Consolidated tools into 5 cohesive modules (`document_tools`, `content_tools`, `format_tools`, `batch_tools`, `protection_tools`).
- **Zero Circular Imports**: Clean dependency tree using PEP 562 lazy loading in `tools/__init__.py`.
- **Backward-Compatible Shims**: `extended_document_tools`, `header_footer_tools`, `link_tools`, and `property_tools` preserved as zero-cost delegators to ensure 100% backward compatibility for existing MCP clients.
- **Parameter Alias Tolerance**: Validation decorators and tools accept either `filename` or `file_path` transparently.
- **Unified Configuration**: Centralized all project metadata and tools (`ruff`, `pytest`, `mypy`) into `pyproject.toml`.

### Removed
- Removed legacy monolithic skill `skills/mcp-word-document-workflow/` (all 23 files).
- Removed redundant configuration files (`ruff.toml`, `pytest.ini`, `mypy.ini`).

---

## [1.1.2] - 2026-03-21

### Added
- **Document Validation Framework**: Comprehensive document validation decorators to verify DOCX file existence, type, and writability
- **Extended Word Document Tools**: Advanced document operations including pagination, complex formatting, and bulk operations
- **Comprehensive Tool Suite**: Document manipulation tools, extended tools, footnote tools, format tools, header/footer tools, link tools, property tools, protection tools with full error handling
- **Server Infrastructure**: Complete MCP server setup with tool registration and context management
- **Response Models**: Structured response models for all tool outputs ensuring consistent client integration

### Changed
- **Project Structure**: Reorganized core modules for better separation of concerns
- **Testing Infrastructure**: Enhanced test suite covering all tool categories and core functionality
- **Manifest Configuration**: Updated MCPBuilder manifest version and configuration

### Fixed
- Document context management and error handling improvements
- Tool response formatting for better client compatibility

---

## [1.1.1] - 2026-03-20

### Added
- **Enhanced Documentation**: Added comprehensive TOOLS.md documentation for all available tools and capabilities

### Changed
- **Docker Support**: Improved Docker configuration and documentation for containerized deployments
- **Dependency Management**: Updated dependencies for better compatibility and performance

### Fixed
- Various bug fixes in document manipulation logic
- Improved error messages for better debugging

---

## [1.1.0] - 2026-03-19

### Added
- **Core Word Document Operations**: Implemented comprehensive document management and manipulation capabilities
- **Style Management**: Enhanced style handling with support for custom styles and formatting
- **Testing Framework**: Added comprehensive test suite for all modules
- **Validation Decorators**: New validation system for document operations
- **Documentation**: Extensive inline documentation and type hints throughout codebase

### Changed
- **Import Organization**: Refactored and cleaned up import statements across all modules for better code organization
- **Module Structure**: Reorganized package structure for improved maintainability
- **Package Version**: Updated from 1.0.1 to 1.1.0

### Fixed
- Import conflicts and circular dependencies
- Module initialization issues

---

## [1.0.1] - 2026-03-15

### Added
- **Docker Support**: Added Dockerfile and .dockerignore for containerized deployments
- **Docker Documentation**: Comprehensive Docker setup and deployment documentation in TOOLS.md

### Changed
- **Package Rename**: Renamed package from `mcp-office-word` to `mcp-word` for better naming consistency
- **Script Updates**: Updated all script and entry-point names to reflect new package name
- **Documentation**: Updated README.md to support additional MCP-compatible clients and improved clarity
- **Manifest Configuration**: Updated manifest.json and configuration files for new package name

### Fixed
- Corrected version numbers in pyproject.toml and __init__.py to 1.0.1
- Fixed trailing spaces and formatting issues in documentation
- Corrected IDE-specific files in .gitignore (.kiro/, .windsurf/)

### Security
- Added proper .gitignore patterns for sensitive and build-related files

---

## [1.0.0] - 2026-02-15

### Added
- **Initial Release**: Complete MCP server implementation for Word document manipulation
- **Document Operations**: Create, read, copy, merge, and convert documents to PDF
- **Content Management**: Add and manipulate headings, paragraphs, tables, pictures, page breaks
- **Search & Replace**: Find and replace content within documents
- **Text Formatting**: Apply comprehensive text formatting including bold, italic, underline, colors
- **Custom Styles**: Create and apply custom document styles
- **Footnotes & Endnotes**: Add, manage, and convert between footnotes and endnotes
- **Security Features**: Password protection, restricted editing mode, digital signature support
- **Headers & Footers**: Create and manage document headers and footers for specific sections
- **Links & Bookmarks**: Insert hyperlinks and internal bookmarks for navigation
- **Document Properties**: Manage core metadata (author, title) and section layouts
- **Python SDK**: Full Python-based MCP server implementation
- **Docker Support**: Dockerfile and documentation for containerized deployment
- **Comprehensive Testing**: Unit tests for core functionality and utilities
- **Documentation**: Complete README with features, installation, and usage instructions

### Dependencies
- `python-docx>=1.0.0`: Core Word document manipulation
- `docx2pdf>=0.1.8`: PDF conversion support
- `msoffcrypto-tool>=5.4.2`: Document protection and encryption
- `fastmcp>=0.9.0`: MCP server framework (FastMCP)

---

## Links

- [GitHub Repository](https://github.com/LuiccianDev/word-mcp)
- [GitHub Issues](https://github.com/LuiccianDev/word-mcp/issues)
- [MCP Protocol](https://modelcontextprotocol.io)
- [Python-docx Documentation](https://python-docx.readthedocs.io)
