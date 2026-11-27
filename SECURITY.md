# Security Policy for Word MCP (`word-mcp`)

Document integrity, sandbox isolation, and host security are fundamental priorities for **Word MCP Server** (`word-mcp`, PyPI: [`luiccian-word-mcp`](https://pypi.org/project/luiccian-word-mcp/)).

Because this Model Context Protocol server interacts with local filesystems, generates documents, and integrates with LLM agents (Claude, Cursor, Cline, Copilot), strict security controls and defense-in-depth principles are enforced across all modules.

---

## Supported Versions

Only the latest major version branch receives active security updates and vulnerability patches.

| Version | Supported | Notes |
| :--- | :--- | :--- |
| **3.x.x** | Yes | Current major branch (latest release: `3.2.0`) |
| **2.x.x** | No | Superseded by 3.0.0; prefixed tool aliases were removed |
| **1.x.x** | No | Deprecated legacy `mcp_word` release |

---

## Security Architecture & Defenses

### 1. Directory Sandboxing (`WORD_MCP_DIRS`)
By default (`./documents`) or via configuration, Word MCP limits document reading, writing, listing, image insertion, and PDF conversions strictly to explicit allowed directories:
- `resolve_sandboxed_path` and `@tool_boundary` validate all primary and secondary path arguments (`filename`, `source_filename`, `destination_filename`, `target_filename`, `source_filenames`, `output_filename`, `image_path`, `directory`, and batch `image` elements) before any filesystem check or directory creation occurs.
- Empty/whitespace paths, NUL bytes (`\x00`), Win32 device prefixes (`\\.\`, `\\?\`), unauthorized UNC shares, and Windows reserved DOS device names (`CON`, `PRN`, `AUX`, `NUL`, `COM1`-`COM9`, `LPT1`-`LPT9`) are rejected before `pathlib.Path.resolve()`.
- Directory traversal attempts (`../`, `..\\`) and symlink escapes outside `WORD_MCP_DIRS` are blocked with `SecurityError` (`error_type="security_violation"`, `recoverable=False`) before any file handle is opened. `list_available_documents` skips symlinks and non-files.
- `_fast_check_ooxml_zip` enforces ZIP entry count, per-entry uncompressed size, total uncompressed size, and compression ratio limits before opening `docx.Document` to mitigate ZIP bombs.

### 2. Password Protection and Protection Metadata
- `protect_document` encrypts the OOXML container with `msoffcrypto-tool` and replaces the target file atomically via `atomic_write_target`. Empty passwords are rejected, and the password is not stored in the document or in this repository.
- `unprotect_document` decrypts the container atomically and returns `Incorrect password` (`error_type="validation"`) when the key is invalid.
- `add_restricted_editing` and `add_digital_signature` write a sidecar JSON file named `<document>.protection` next to the `.docx`. Sidecar reads/deletes enforce `WORD_MCP_DIRS` containment, reject symlinks, cap metadata size at 64 KiB, and bound PBKDF2-HMAC-SHA-256 iteration counts. The `.protection` file is advisory metadata for MCP clients and does not replace native OOXML encryption (`protect_document`).
- `add_digital_signature` writes a visible signature block and stores a SHA-256 hash of body paragraphs, table cells, and headers/footers. It detects content or signature-block changes after signing. It is **not a cryptographic signature** and does not verify signer identity.
- `verify_document` checks `.protection` sidecar metadata (and invokes `core_verify_signature` for signed documents), or falls back to `msoffcrypto` container inspection and password verification for password-encrypted `.docx` files. Passing an empty password (`""`) returns verification failure. It does not validate cryptographic OOXML package signatures.

### 3. Safe Subprocess, Link, and Transport Controls
- Subprocess invocations (such as LibreOffice headless PDF conversion) use argument lists directly rather than `shell=True`, stage output inside a private `TemporaryDirectory`, and enforce execution timeouts.
- `core_add_hyperlink` and inline markdown links allow only `http://`, `https://`, and `mailto:` URLs.
- The server communicates exclusively over local `stdio` and opens no network ports.

### 4. Known Limitations
- The `.protection` sidecar is advisory metadata. Microsoft Word does not enforce the `editable_sections` list.
- Page count is not computed by `get_document_info`; rendering is required for an authoritative value.
- PDF conversion depends on LibreOffice on Linux and `docx2pdf` on Windows/macOS. Failure is reported as a structured error.

---

## Reporting a Vulnerability

If you discover a security vulnerability in **Word MCP Server**, please report it responsibly:

1. **Private Disclosure**: Do **not** open a public issue or discussion on GitHub.
2. **Contact**: Send an email directly to the maintainer at [luicciandev@gmail.com](mailto:luicciandev@gmail.com) with the subject:
   `[SECURITY VULNERABILITY] Word MCP - <Brief Description>`
3. **Information to Include**:
   - Detailed description of the vulnerability and its potential impact.
   - Proof-of-concept (PoC) code or exact steps to reproduce.
   - Affected version(s) and operating system.
   - Any suggested mitigations or patches if available.

### Response Timeline
- **Initial Acknowledgment**: Within 48 hours.
- **Triage & Assessment**: Within 5 business days.
- **Fix & Advisory**: A patched release will be published via GitHub Releases and PyPI with coordinated disclosure.
