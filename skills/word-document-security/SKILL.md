---
name: word-document-security
description: "Protect Word documents with passwords, configure editing restrictions, verify file integrity, and export to PDF for final distribution using Word MCP."
argument-hint: "Target document path, protection mode (read-only, comments, tracked changes), password, or export target"
user-invocable: true
disable-model-invocation: false
---

# Word Document Security and Distribution

Use this skill to protect Microsoft Word documents, enforce editing permissions, verify structural integrity, and export files to PDF format.

## Core Rules

1. **Document Integrity Verification:**
   - Run `word_protection(action="verify", ...)` (or `verify_document` in `full` mode) to inspect sidecar protection metadata or `msoffcrypto` encryption status before distribution.
2. **Protection Modes & Atomic Persistence:**
   - **Container encryption**: `protect_document` (`word_protection` `action="protect"`) encrypts the `.docx` with a password using `msoffcrypto` and replaces the target file atomically.
   - **Restricted editing metadata**: `add_restricted_editing` (`word_protection` `action="restrict_editing"`) records an advisory list of editable sections in the `.protection` sidecar. Microsoft Word does not enforce this list.
   - **Signature marker**: `add_digital_signature` (`word_protection` `action="add_signature"`) writes a visible signature block and a SHA-256 content hash. It is not a cryptographic signature.
3. **Distribution & PDF Conversion:**
   - For finalized documents intended for signing or non-editable distribution, export via `word_document(action="to_pdf", ...)` (or `convert_to_pdf`). Both `filename` and `output_filename` must reside within `WORD_MCP_DIRS`.

## Usage Patterns

### Pattern A: Lock Document with Password
```python
# Compound mode (default):
word_protection(action="protect", params={"filename": "/path/to/contrato_final.docx", "password": "SecurePassword2026!"})

# Full mode equivalent:
protect_document(
    filename="/path/to/contrato_final.docx",
    password="SecurePassword2026!"
)
```

### Pattern B: Unlock Protected Document
```python
# Compound mode (default):
word_protection(action="unprotect", params={"filename": "/path/to/contrato_final.docx", "password": "SecurePassword2026!"})

# Full mode equivalent:
unprotect_document(
    filename="/path/to/contrato_final.docx",
    password="SecurePassword2026!"
)
```

### Pattern C: Restrict Editing Sections
```python
add_restricted_editing(
    filename="/path/to/borrador_revisores.docx",
    password="SecurePassword2026!",
    editable_sections=["Introduccion", "Conclusiones"]
)
```

### Pattern D: Export to PDF for Distribution
```python
# Compound mode (default):
word_document(action="to_pdf", params={"filename": "/path/to/informe_anual.docx", "output_filename": "/path/to/informe_anual.pdf"})

# Full mode equivalent:
convert_to_pdf(
    filename="/path/to/informe_anual.docx",
    output_filename="/path/to/informe_anual.pdf"
)
```
