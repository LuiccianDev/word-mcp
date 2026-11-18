# Document Editing Tools Reference

This reference describes the primary modification and inspection tools available in `word-mcp`.

## Inspection Tools (`word_document` in compound mode)

| Tool / Compound Action | Tokens | Purpose |
| :--- | :--- | :--- |
| `get_document_outline` (`action="outline"`) | Ultra-Low (~100 tokens) | Extracts paragraph previews, styles, and table dimensions without loading full body text. |
| `read_document_range` (`action="read_range"`) | Low (bounded by `max_chars`) | Reads a slice of paragraphs by index or `section_heading` without loading the entire document. |
| `get_document_tables` (`action="get_tables"`) | Low (paginated rows) | Inspects table metadata (`row_count`, `col_count`, `headers`) and optionally paginates rows. |
| `get_document_info` (`action="info"`) | Very Low (~150 tokens) | Metadata (author, title, modified date, size_kb) and core document properties. |
| `find_text_in_document` (`action="find_text"`) | Low (~200 tokens) | Searches a query substring and returns matching paragraph indexes and context snippets. |
| `get_document_text` (`action="text"`) | High (proportional to doc) | Dumps complete text. Use only when full semantic text analysis is strictly needed. |

## Mutation & Batch Tools (`batch_add_content` & `word_content` in compound mode)

| Tool / Compound Action | Best Used For |
| :--- | :--- |
| `batch_add_content` | Adding multi-element sections (headings, paragraphs, tables, callouts, lists) in a single atomic MCP request. |
| `batch_edit_document` (`action="batch_edit"`) | Applying multiple `replace_text`, `replace_paragraph`, `insert_after_paragraph`, and `delete_paragraph` operations atomically in one pass. |
| `search_and_replace` (`action="search_and_replace"`) | Global string replacement across all paragraphs and tables. |
| `merge_documents` (`word_document` `action="merge"`) | Concatenating multiple sandboxed `.docx` files with optional page breaks between them. |
| `add_hyperlink` (`action="add_hyperlink"`) | Inserting an external URL or internal bookmark link into a paragraph. |
| `add_bookmark` (`action="add_bookmark"`) | Creating an internal navigation bookmark for cross-references. |
| `add_footnote_to_document` (`word_footnote` `action="add_footnote"`) | Adding academic or legal citations linked to a specific paragraph. |
