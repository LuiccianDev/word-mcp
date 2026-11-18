---
name: word-document-editor
description: "Inspect, modify, search/replace, append batch content, or merge existing Word documents (.docx) efficiently using Word MCP. Prioritizes lightweight outline inspection and batch updates over chatty micro-calls."
argument-hint: "Target document path, edit action (inspect, search/replace, append batch content, merge), and modification parameters"
user-invocable: true
disable-model-invocation: false
---

# Word Document Editor

Use this skill to inspect and update **existing** Microsoft Word documents with surgical precision and minimal token consumption.

## Core Rules

1. **Inspect Before Reading Full Text:**
   - Use `word_document(action="outline", ...)` (or `get_document_outline` in `full` mode) first to see paragraph previews, paragraph styles, and table dimensions.
   - Use `word_document(action="read_range", ...)` or `word_document(action="get_tables", ...)` to read targeted sections or paginated table rows.
   - Do NOT call `get_document_text` (`word_document(action="text", ...)`) unless you strictly need to read the full body content.
2. **Batch Appending & Atomic Edits Over Single Injections:**
   - When adding new sections, appendices, or summary tables, use `batch_add_content`.
   - When applying multiple text/paragraph replacements or insertions in one pass, use `word_content(action="batch_edit", ...)`.
   - Never chain single `add_heading` + multiple `add_paragraph` + `add_table` calls.
3. **Surgical Search and Replace:**
   - Use `word_content(action="search_and_replace", ...)` (or `search_and_replace`) for template variables (e.g. `{{CLIENT_NAME}}`, `[FECHA]`) or updating outdated terms across the entire document.
4. **Document Merging:**
   - When assembling multiple chapters or reports, use `word_document(action="merge", ...)` (or `merge_documents`) rather than manually copying text between files. Both `target_filename` and `source_filenames` must reside within `WORD_MCP_DIRS`.

## Workflow Patterns

### Pattern A: Inspect Structure & Targeted Sections
```python
# Compound mode (default): inspect structure or read a bounded section
word_document(action="outline", params={"filename": "/path/to/documento.docx"})
word_document(action="read_range", params={"filename": "/path/to/documento.docx", "section_heading": "Introduccion"})

# Full mode equivalent:
get_document_outline(filename="/path/to/documento.docx")
find_text_in_document(filename="/path/to/documento.docx", text_to_find="clausula")
```

### Pattern B: Search & Replace or Multi-Op Batch Edit
```python
# Compound mode (default): atomic multi-operation batch edit
word_content(
    action="batch_edit",
    params={
        "filename": "/path/to/contrato.docx",
        "operations": [
            {"op": "replace_text", "old_text": "ACME Corp", "new_text": "Global Logistics Inc."},
            {"op": "replace_paragraph", "paragraph_index": 2, "text": "Cláusula actualizada."}
        ]
    }
)

# Full mode single search_and_replace:
search_and_replace(
    filename="/path/to/contrato.docx",
    find_text="ACME Corp",
    replace_text="Global Logistics Inc."
)
```

### Pattern C: Append a Full New Section in 1 Call
```python
# Use batch_add_content to add a heading, paragraph, and table simultaneously
batch_add_content(
    filename="/path/to/informe.docx",
    elements=[
        {
            "type": "heading",
            "text": "Anexo A: Datos Complementarios",
            "level": 2
        },
        {
            "type": "paragraph",
            "text": "La siguiente tabla resume los indicadores complementarios auditados:"
        },
        {
            "type": "table",
            "headers": ["Métrica", "2025", "2026"],
            "rows": [
                ["Eficiencia", "84%", "92%"],
                ["Disponibilidad", "99.1%", "99.9%"]
            ]
        },
        {
            "type": "page_break"
        }
    ]
)
```

### Pattern D: Merge Documents
```python
# Merge appendices or chapters into the main document
merge_documents(
    target_filename="/path/to/informe_final.docx",
    source_filenames=["/path/to/anexo1.docx", "/path/to/anexo2.docx"],
    add_page_breaks=True
)
```
