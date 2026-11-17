# Context-Efficient Inspection, Batch Editing, and Single-Pass Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add context-bounded document inspection (`read_range`, `get_tables`), single-pass multi-edit (`batch_edit`), and richer single-pass generation (`image` element, `callout` `title`, inline `[text](url)` hyperlinks) without increasing the MCP tool count in `compound` or `granular` mode.

**Architecture:** New inspection and editing capabilities are implemented as pure helpers in `word_mcp.core`, wrapped as async tool functions in `word_mcp.tools`, and exposed as new actions on the existing `word_document` and `word_content` compound tools (`WORD_MCP_TOOL_MODE=compound` remains 9 tools; `granular` mode remains 40 tools). Single-pass enhancements live directly in `word_mcp.core.batch` so `build_document`, `create_from_markdown`, and `batch_add_content` gain them automatically with zero schema bloat.

**Tech Stack:** Python 3.10+, `python-docx`, `fastmcp`, `pydantic`, `pytest`, `ruff`, `mypy`.

**Spec:** User-approved planning scope (2026-11-16): context-efficient section/table reading, atomic `batch_edit`, and single-pass `image` / `callout` `title` / inline hyperlink support integrated into existing compound & single-pass tools.

## Global Constraints

- Python 3.10+ compatibility; all `.py` files must use `snake_case.py` naming and pass `ruff` and `mypy`.
- `WORD_MCP_TOOL_MODE=compound` MUST continue to register exactly 9 tools (`EXPECTED_COMPOUND_TOOLS` in `tests/test_compound_tools.py`).
- `WORD_MCP_TOOL_MODE=granular` (default) MUST continue to register exactly 40 tools (`test_registers_exactly_forty_tools` in `tests/test_docs_tool_parity.py`).
- All repository code comments, docstrings, documentation, and commit messages MUST be in Technical English with zero informal emojis.
- Read/inspection actions MUST bound output size (`max_chars`, `row_limit`, `include_rows`) to protect LLM context windows.
- All file paths (including `image` paths in `batch.py`) MUST respect `MCP_ALLOWED_DIRECTORIES` via `is_path_in_allowed_directories`.

## Review Focus

- `read_range` with an empty document, out-of-range `start_index`/`end_index`, or non-existent `section_heading`: must return a structured result (`paragraphs=[]` or clear error) and enforce `max_chars` truncation with `truncated=True` and `next_index`.
- `get_tables` on a document with 0 tables, out-of-range `table_index`, or large tables (100+ rows): must default to metadata-only (`include_rows=False`) and cap `row_limit` so large tables never flood context.
- `batch_edit` with multiple `delete_paragraph` operations or invalid operation payloads: must validate all operations before mutating/saving and handle index ordering deterministically without shifting bugs.
- `image` element in `build_document`/`batch_add_content` with missing file, path outside `MCP_ALLOWED_DIRECTORIES`, or `width <= 0`: must raise `FileNotFoundError`/`ValueError` cleanly before saving.
- Inline `[text](url)` parsing in `add_inline_formatted_text` with malformed brackets, empty URL, or adjacent `**bold**`/`*italic*` tokens: must keep malformed syntax literal and only emit `w:hyperlink` for valid `[label](url)` / `[label](#bookmark)` spans.

---

### Task 1: Context-Bounded Range & Section Reading (`read_range`)

**Files:**
- Modify: `src/word_mcp/core/document.py`
- Modify: `src/word_mcp/tools/document_tools.py`
- Modify: `src/word_mcp/tools/compound_tools.py:255-266`
- Test: `tests/test_document_tools.py`, `tests/test_compound_tools.py`

**Interfaces:**
- Produces: `core_read_range(doc: DocumentType, start_index: int = 0, end_index: int | None = None, section_heading: str | None = None, max_chars: int = 4000) -> dict[str, Any]`
- Produces: `async def read_document_range(filename: str, start_index: int = 0, end_index: int | None = None, section_heading: str | None = None, max_chars: int = 4000) -> dict[str, Any]`
- Produces: `word_document(action="read_range", params={...})`

- [ ] **Step 1: Write the failing tests**

```python
# In tests/test_document_tools.py
@pytest.mark.asyncio
async def test_read_document_range_by_section_and_max_chars(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "sample.docx"
    # Build doc with Heading 1 "Intro", 2 paragraphs, Heading 1 "Details", 2 long paragraphs
    result = await read_document_range(str(doc_path), section_heading="Intro", max_chars=200)
    assert result["status"] == "success"
    assert len(result["paragraphs"]) >= 1
    assert all(p["index"] < 3 for p in result["paragraphs"])

@pytest.mark.asyncio
async def test_read_document_range_truncates_on_max_chars(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "large.docx"
    # Build doc with 10 paragraphs of 100 chars each
    result = await read_document_range(str(doc_path), start_index=0, max_chars=250)
    assert result["status"] == "success"
    assert result["truncated"] is True
    assert result["next_index"] == result["paragraphs"][-1]["index"] + 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_document_tools.py -k "read_document_range" -v`
Expected: FAIL with `ImportError` / `NameError` for `read_document_range`.

- [ ] **Step 3: Implement `core_read_range` in `src/word_mcp/core/document.py`**

Implement `core_read_range(doc: DocumentType, start_index: int = 0, end_index: int | None = None, section_heading: str | None = None, max_chars: int = 4000) -> dict[str, Any]`.
When `section_heading` is provided, locate the first heading paragraph (`style.name` starting with `"Heading"`) whose text contains `section_heading` (case-insensitive), record its level, and slice until the next heading of equal or higher level (smaller/equal integer); if not found, return `paragraphs=[]`, `section_found=False`, `truncated=False`, `next_index=None`. Enforce `max_chars >= 1` (accumulate character length of returned paragraph texts; always include at least the first paragraph if `start_index` is in range, or stop before exceeding `max_chars` and set `truncated=True`, `next_index=idx`).

- [ ] **Step 4: Implement `read_document_range` in `src/word_mcp/tools/document_tools.py` and register `"read_range"` in `src/word_mcp/tools/compound_tools.py`**

Decorate `read_document_range` with `@validate_docx_read("filename")`, open via `document_context(filename, mode="read")`, call `core_read_range`, and wrap errors with `ExceptionTool.handle_error`. Add `"read_range": document_tools.read_document_range` to `_DOCUMENT_ACTIONS` in `compound_tools.py`.

- [ ] **Step 5: Add compound test in `tests/test_compound_tools.py` and run tests**

Add a test calling `word_document` with `action="read_range"` and `action="help", params={"action": "read_range"}`.
Run: `uv run pytest tests/test_document_tools.py tests/test_compound_tools.py -k "read_range or read_document_range" -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/word_mcp/core/document.py src/word_mcp/tools/document_tools.py src/word_mcp/tools/compound_tools.py tests/test_document_tools.py tests/test_compound_tools.py
git commit -m "feat: add context-bounded read_range action to word_document"
```

---

### Task 2: Context-Bounded Table Inspection (`get_tables`)

**Files:**
- Modify: `src/word_mcp/core/tables.py`
- Modify: `src/word_mcp/tools/document_tools.py`
- Modify: `src/word_mcp/tools/compound_tools.py:255-266`
- Test: `tests/test_document_tools.py`, `tests/test_compound_tools.py`

**Interfaces:**
- Produces: `core_get_tables(doc: DocumentType, table_index: int | None = None, include_rows: bool = False, row_offset: int = 0, row_limit: int = 20) -> dict[str, Any]`
- Produces: `async def get_document_tables(filename: str, table_index: int | None = None, include_rows: bool = False, row_offset: int = 0, row_limit: int = 20) -> dict[str, Any]`
- Produces: `word_document(action="get_tables", params={...})`

- [ ] **Step 1: Write the failing tests**

```python
@pytest.mark.asyncio
async def test_get_document_tables_metadata_and_paginated_rows(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "tables.docx"
    # Build doc with a 5-row table
    meta = await get_document_tables(str(doc_path))
    assert meta["status"] == "success"
    assert meta["total_tables"] == 1
    assert "rows" not in meta["tables"][0]

    page = await get_document_tables(str(doc_path), table_index=0, include_rows=True, row_offset=1, row_limit=2)
    assert len(page["tables"][0]["rows"]) == 2
    assert page["tables"][0]["truncated"] is True
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_document_tools.py -k "get_document_tables" -v`
Expected: FAIL with `ImportError` / `NameError`.

- [ ] **Step 3: Implement `core_get_tables` in `src/word_mcp/core/tables.py`**

Implement `core_get_tables(doc: DocumentType, table_index: int | None = None, include_rows: bool = False, row_offset: int = 0, row_limit: int = 20) -> dict[str, Any]`.
Clamp `row_offset >= 0` and `1 <= row_limit <= 100`. If `table_index is not None`, validate `0 <= table_index < len(doc.tables)` (raising `IndexError` if out of range) and inspect only that table; otherwise inspect all tables. Each table entry contains `table_index`, `row_count`, `col_count`, and `headers` (cell texts of row 0 if present). Only include `rows` (sliced `[row_offset : row_offset + row_limit]`), `row_offset`, and `truncated` when `include_rows=True`.

- [ ] **Step 4: Implement `get_document_tables` in `src/word_mcp/tools/document_tools.py` and register `"get_tables"` in `compound_tools.py`**

Decorate with `@validate_docx_read("filename")`, open via `document_context(filename, mode="read")`, and add `"get_tables": document_tools.get_document_tables` to `_DOCUMENT_ACTIONS`.

- [ ] **Step 5: Run tests and verify pass**

Run: `uv run pytest tests/test_document_tools.py tests/test_compound_tools.py -k "tables" -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/word_mcp/core/tables.py src/word_mcp/tools/document_tools.py src/word_mcp/tools/compound_tools.py tests/test_document_tools.py tests/test_compound_tools.py
git commit -m "feat: add paginated get_tables action to word_document"
```

---

### Task 3: Atomic Multi-Operation `batch_edit` (`word_content` action `"batch_edit"`)

**Files:**
- Modify: `src/word_mcp/core/content.py`
- Modify: `src/word_mcp/tools/content_tools.py`
- Modify: `src/word_mcp/tools/compound_tools.py:268-279`
- Test: `tests/test_content_tools.py`, `tests/test_compound_tools.py`

**Interfaces:**
- Produces: `core_batch_edit(doc: DocumentType, operations: list[dict[str, Any]]) -> dict[str, Any]`
- Produces: `async def batch_edit_document(filename: str, operations: list[dict[str, Any]]) -> dict[str, Any]`
- Produces: `word_content(action="batch_edit", params={"filename": ..., "operations": [...]})`

- [ ] **Step 1: Write the failing tests**

```python
@pytest.mark.asyncio
async def test_batch_edit_document_applies_multiple_ops_and_rejects_invalid(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    doc_path = tmp_path / "edit.docx"
    # Create doc with 3 paragraphs: "Alpha", "Beta", "Gamma"
    res = await batch_edit_document(
        str(doc_path),
        operations=[
            {"op": "replace_text", "old_text": "Alpha", "new_text": "Alpha One"},
            {"op": "replace_paragraph", "paragraph_index": 1, "text": "Beta Two"},
            {"op": "delete_paragraph", "paragraph_index": 2},
        ],
    )
    assert res["status"] == "success"
    assert res["applied_count"] == 3
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_content_tools.py -k "batch_edit" -v`
Expected: FAIL with `ImportError` / `NameError`.

- [ ] **Step 3: Implement `core_batch_edit` in `src/word_mcp/core/content.py`**

Supported `op` values:
- `"replace_text"`: `{"op": "replace_text", "old_text": str, "new_text": str}` -> delegates to `core_find_and_replace_text`.
- `"replace_paragraph"`: `{"op": "replace_paragraph", "paragraph_index": int, "text": str}` -> validates index and updates `doc.paragraphs[idx].text`.
- `"delete_paragraph"`: `{"op": "delete_paragraph", "paragraph_index": int}` -> validates index and removes XML element.
- `"insert_after_paragraph"`: `{"op": "insert_after_paragraph", "paragraph_index": int, "text": str, "style": str | None}` -> inserts a new paragraph XML element immediately after `doc.paragraphs[idx]._p` using `OxmlElement("w:p")` / `addnext` or `insert_paragraph_before` on `idx + 1`.
Validate all `op` names and required keys upfront before mutating `doc` so invalid batches fail fast without partial writes.

- [ ] **Step 4: Implement `batch_edit_document` in `src/word_mcp/tools/content_tools.py` and register `"batch_edit"` in `compound_tools.py`**

Decorate with `@validate_docx_write("filename")`, run inside `document_context(filename, mode="write")`, and register `"batch_edit": content_tools.batch_edit_document` in `_CONTENT_ACTIONS`.

- [ ] **Step 5: Run tests and verify pass**

Run: `uv run pytest tests/test_content_tools.py tests/test_compound_tools.py -k "batch_edit" -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/word_mcp/core/content.py src/word_mcp/tools/content_tools.py src/word_mcp/tools/compound_tools.py tests/test_content_tools.py tests/test_compound_tools.py
git commit -m "feat: add atomic batch_edit action to word_content"
```

---

### Task 4: Single-Pass `image` Element & `callout` Title in `build_document` / `batch_add_content`

**Files:**
- Modify: `src/word_mcp/core/batch.py:36-48, 295-306`
- Test: `tests/test_batch_tools.py`

**Interfaces:**
- Modifies: `SUPPORTED_ELEMENT_TYPES` in `src/word_mcp/core/batch.py` (adds `"image"`)
- Modifies: `render_element(doc: DocumentType, elem: dict[str, Any]) -> str` in `src/word_mcp/core/batch.py`

- [ ] **Step 1: Write the failing tests**

```python
@pytest.mark.asyncio
async def test_build_document_supports_image_and_callout_title(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_ALLOWED_DIRECTORIES", str(tmp_path))
    img_path = tmp_path / "pixel.png"
    # Write minimal valid 1x1 PNG bytes
    doc_path = tmp_path / "with_image.docx"
    res = await build_document(
        filename=str(doc_path),
        elements=[
            {"type": "callout", "title": "NOTE", "text": "Important detail"},
            {"type": "image", "path": str(img_path), "width": 2.0, "caption": "Figure 1", "alignment": "center"},
        ],
    )
    assert res["status"] == "success"
    assert res["element_counts"]["image"] == 1
    assert res["element_counts"]["callout"] == 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_batch_tools.py -k "image_and_callout_title" -v`
Expected: FAIL with `ValueError: Unknown element type 'image'`.

- [ ] **Step 3: Implement `image` rendering and `callout` `title` in `src/word_mcp/core/batch.py`**

- Add `"image"` to `SUPPORTED_ELEMENT_TYPES`.
- In `render_element`, handle `elem_type in ("image", "picture")`: read `image_path = elem.get("path") or elem.get("image_path")`, validate non-empty string, `is_path_in_allowed_directories(abs_path)`, `os.path.exists(abs_path)`, and `width > 0` if `width is not None`. Add a paragraph (with optional `alignment`), add a run with `run.add_picture(abs_path, width=Inches(float(width)))` (or without width), and if `caption` is provided add a centered/aligned italic caption paragraph below. Return `"image"`.
- In `callout` branch, if `title := elem.get("title")`, prepend a bold run `f"{title}: "` before `add_inline_formatted_text(p, text)`.

- [ ] **Step 4: Run tests and verify pass**

Run: `uv run pytest tests/test_batch_tools.py -k "image or callout" -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/core/batch.py tests/test_batch_tools.py
git commit -m "feat: support image elements and callout titles in single-pass batch builder"
```

---

### Task 5: Inline `[text](url)` Hyperlinks & Documentation Sync

**Files:**
- Modify: `src/word_mcp/core/batch.py:139-160`
- Modify: `docs/TOOLS.md`
- Test: `tests/test_batch_tools.py`, `tests/test_docs_tool_parity.py`, `tests/test_docs_content_contract.py`

**Interfaces:**
- Modifies: `add_inline_formatted_text(paragraph: Any, text: str) -> None` in `src/word_mcp/core/batch.py`

- [ ] **Step 1: Write the failing tests**

```python
def test_add_inline_formatted_text_renders_markdown_hyperlinks():
    doc = Document()
    p = doc.add_paragraph()
    add_inline_formatted_text(p, "Visit [Word MCP](https://example.com) for **docs**.")
    hyperlinks = p._p.findall(qn("w:hyperlink"))
    assert len(hyperlinks) == 1
    assert "Visit " in p.text
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_batch_tools.py -k "hyperlinks" -v`
Expected: FAIL (`len(hyperlinks) == 0`).

- [ ] **Step 3: Implement inline `[label](url)` parsing in `add_inline_formatted_text` in `src/word_mcp/core/batch.py`**

Update `add_inline_formatted_text` to split on `(\[[^\[\]]+\]\([^\s()]+\)|\*\*.*?\*\*|\*.*?\*)`. For `[label](target)` tokens: if `target` starts with `#` (internal bookmark, length > 1), emit a `w:hyperlink` with `w:anchor`; if `target` starts with `http://`, `https://`, or `mailto:`, emit a `w:hyperlink` with `part.relate_to(target, RT.HYPERLINK, is_external=True)` and `r:id`; otherwise leave literal text.

- [ ] **Step 4: Update `docs/TOOLS.md` and run full verification suite**

Document `read_range`, `get_tables`, `batch_edit` in the Compound Mode / Tool Reference sections, document the `image` element and `callout` `title` and inline `[text](url)` links under `build_document` / `create_from_markdown` / `batch_add_content` in `docs/TOOLS.md`.
Run:
- `uv run pytest -q`
- `uv run ruff check src tests scripts`
- `uv run mypy src`
- `uvx --from graphifyy graphify update .`
Expected: All tests pass, zero lint/type errors, tool count remains 40 (granular) / 9 (compound).

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/core/batch.py docs/TOOLS.md tests/test_batch_tools.py
git commit -m "feat: support inline markdown hyperlinks in batch builder and update docs"
```
