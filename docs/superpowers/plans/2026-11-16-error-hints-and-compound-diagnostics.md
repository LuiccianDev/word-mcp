# Error Diagnostics, Type Hints, and Compound Tool Guidance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix exception-mapping inheritance shadowing in `ExceptionTool`, add actionable `hint` guidance to document validators, and upgrade `compound_tools` with fuzzy suggestions, action schema inspection (`help`), and parameter/type diagnostics.

**Architecture:** Three focused, independent improvements: (1) reorder `ExceptionTool.EXCEPTION_MAPPING` so specific `OSError` subclasses (`FileNotFoundError`, `PermissionError`) match before `OSError`; (2) enrich `_validate_docx_path` and `_check_file_writeable` error payloads in `document_validators.py` with a non-breaking `"hint"` key while preserving `"error"`; and (3) upgrade `_dispatch` and `_Params` in `compound_tools.py` to emit `difflib` close-match hints, expected parameter signatures, type-mismatch diagnostics, and a built-in `"help"` action per compound tool.

**Tech Stack:** Python 3.11+, FastMCP (`ToolError`), Pydantic v2, `inspect`, `difflib`, `pytest`, `ruff`, `mypy`.

**Spec:** User request in session `b8a007a8-a814-4eab-ade3-95d0b95bb3d9` targeting bug fixes, error message clarity, and accurate type/hint diagnostics across `word-mcp`.

## Global Constraints

- Python >=3.11; all source modules in `src/word_mcp/` must pass `uv run ruff check src tests scripts` and `uv run mypy src`.
- All repository documentation, code comments, and commit messages must be in Technical English with zero informal emojis.
- Preserve existing `"error"` keys in validator dict responses for backwards compatibility with existing callers and tests.
- Run `uv run pytest` after changes to confirm all tests pass without skipped or relaxed assertions.
- After modifying `src/word_mcp/`, update Graphify deterministically (`uvx --from graphifyy graphify update .`).

## Review Focus

- `FileNotFoundError` and `PermissionError` subclassing `OSError`: `ExceptionTool.get_error_info` must classify them as `"file_not_found"` (`recoverable=False`) and `"permission_denied"` (`recoverable=False`), never `"file_operation"`.
- `compound_tools` `"help"` action on `word_document`, `word_content`, `word_format`, `word_protection`, `word_footnote`, `word_property`: must work both with `params=None` (list all actions and summaries) and `params={"action": "<name>"}` (return exact required/optional parameter names, types, and defaults), and reject unknown `params["action"]` with a `ToolError` suggestion.
- `compound_tools` missing/unknown/wrong-type `params`: passing a wrong type (e.g., `"level": "1"` or `"rows": "two"` or `"params": "bad"`) or misspelled action/param (e.g., `"creat"`, `"filname"`) must raise `ToolError` with a `Did you mean '...'?` or expected-type hint instead of leaking a raw Python `TypeError`.
- Validator `"hint"` field on non-existent, wrong-extension, corrupt, or locked `.docx` files: must return `{"error": ..., "hint": ...}` without breaking callers that only read `result["error"]`.
- `compound_tools` `action="help"` must not be shadowed if any domain map ever defines a real `"help"` action, and `WORD_MCP_TOOL_MODE=compound` must still register exactly 9 tools.

---

### Task 1: Fix `OSError` Subclass Shadowing in `ExceptionTool.EXCEPTION_MAPPING`

**Files:**
- Modify: `src/word_mcp/exception/exception_tools.py:23-64`
- Test: `tests/test_exception_centralization.py`

**Interfaces:**
- Consumes: `ExceptionTool.EXCEPTION_MAPPING`, `ExceptionTool.get_error_info(exception: Exception) -> dict[str, Any]`, `ExceptionTool.handle_error(exception: Exception, filename: str | None = None, operation: str | None = None) -> dict[str, Any]`
- Produces: Accurate `error_type` classification (`"file_not_found"`, `"permission_denied"`, `"file_operation"`) and matching `suggestion` strings from `ExceptionTool._get_suggestion`.

- [ ] **Step 1: Write the failing test in `tests/test_exception_centralization.py`**

```python
def test_oserror_subclasses_not_shadowed_by_oserror():
    fnf_info = ExceptionTool.get_error_info(FileNotFoundError("missing.docx"))
    assert fnf_info["error_type"] == "file_not_found"
    assert fnf_info["recoverable"] is False

    perm_info = ExceptionTool.get_error_info(PermissionError("locked.docx"))
    assert perm_info["error_type"] == "permission_denied"
    assert perm_info["recoverable"] is False

    os_info = ExceptionTool.get_error_info(OSError("generic io"))
    assert os_info["error_type"] == "file_operation"
    assert os_info["recoverable"] is True
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_exception_centralization.py -k test_oserror_subclasses_not_shadowed_by_oserror -v`
Expected: FAIL with `AssertionError: assert 'file_operation' == 'file_not_found'`

- [ ] **Step 3: Reorder `ExceptionTool.EXCEPTION_MAPPING` in `src/word_mcp/exception/exception_tools.py`**

Move `FileNotFoundError` and `PermissionError` entries above `IOError` and `OSError` in `EXCEPTION_MAPPING` so `issubclass` matches the most specific exception classes first. Add a brief inline comment explaining why specific `OSError` subclasses must precede `OSError`.

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_exception_centralization.py -v`
Expected: PASS (all tests in `test_exception_centralization.py` pass)

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/exception/exception_tools.py tests/test_exception_centralization.py
git commit -m "fix(exception): place FileNotFoundError and PermissionError before OSError in EXCEPTION_MAPPING"
```

---

### Task 2: Add Actionable `hint` Guidance to Document Validators

**Files:**
- Modify: `src/word_mcp/validation/document_validators.py:163-183` (and `_check_file_writeable` error returns)
- Test: `tests/test_document_validators_hints.py` (new focused test file)

**Interfaces:**
- Consumes: `_validate_docx_path(path_str: str) -> dict[str, str] | None`, `_check_file_writeable(path_str: str) -> tuple[bool, str]`
- Produces: Dict error payloads `{"error": str, "hint": str}` from `validate_docx_file`, `validate_docx_read`, `validate_docx_write`, and `check_file_writeable` decorators.

- [ ] **Step 1: Write the failing tests in `tests/test_document_validators_hints.py`**

```python
from pathlib import Path
from word_mcp.validation.document_validators import _validate_docx_path, validate_docx_read


def test_validate_docx_path_missing_file_includes_hint(tmp_path: Path):
    missing = tmp_path / "missing.docx"
    result = _validate_docx_path(str(missing))
    assert result is not None
    assert "does not exist" in result["error"]
    assert "hint" in result
    assert "create_document" in result["hint"]


def test_validate_docx_path_wrong_extension_includes_hint(tmp_path: Path):
    txt = tmp_path / "notes.txt"
    txt.write_text("hello", encoding="utf-8")
    result = _validate_docx_path(str(txt))
    assert result is not None
    assert "not a .docx" in result["error"]
    assert ".docx" in result["hint"]


def test_validate_docx_path_corrupt_file_includes_hint(tmp_path: Path):
    bad = tmp_path / "corrupt.docx"
    bad.write_bytes(b"not a zip")
    result = _validate_docx_path(str(bad))
    assert result is not None
    assert "not a valid Word document" in result["error"]
    assert "hint" in result


async def test_validate_docx_read_returns_hint_for_missing_file(tmp_path: Path):
    @validate_docx_read("filename")
    async def dummy(filename: str) -> dict[str, str]:
        return {"status": "ok"}

    result = await dummy(filename=str(tmp_path / "nope.docx"))
    assert "error" in result
    assert "hint" in result
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_document_validators_hints.py -v`
Expected: FAIL with `AssertionError: assert 'hint' in {'error': ...}`

- [ ] **Step 3: Add `"hint"` messages in `src/word_mcp/validation/document_validators.py`**

Update `_validate_docx_path(path_str: str) -> dict[str, str] | None` to return `"hint"` alongside `"error"` for:
- missing file: suggest checking the path with `list_available_documents` or creating it first with `create_document` / `build_document`.
- non-`.docx` extension: suggest renaming or using a `.docx` extension (legacy `.doc` files must be converted to `.docx` first).
- `PackageNotFoundError` / corrupt file: suggest verifying the file is a valid Office Open XML `.docx` package and not a renamed text/legacy `.doc` file.
- generic open error: suggest ensuring the file is not locked by Microsoft Word or another process.
Also update `check_file_writeable` / `validate_docx_write` wrapper returns so `{"error": error_msg, "hint": ...}` includes a write-permission hint when `_check_file_writeable` returns `False`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_document_validators_hints.py tests/test_document_tools.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/validation/document_validators.py tests/test_document_validators_hints.py
git commit -m "feat(validation): add actionable hint guidance to document validator errors"
```

---

### Task 3: Add Fuzzy Suggestions, Parameter Schema Hints, Type Diagnostics, and `help` Action to `compound_tools`

**Files:**
- Modify: `src/word_mcp/tools/compound_tools.py:32-100,168-238`
- Test: `tests/test_compound_tools.py`

**Interfaces:**
- Consumes: `_DOCUMENT_ACTIONS`, `_CONTENT_ACTIONS`, `_FORMAT_ACTIONS`, `_PROTECTION_ACTIONS`, `_FOOTNOTE_ACTIONS`, `_PROPERTY_ACTIONS`, `_Params`, `_dispatch`
- Produces:
  - `_format_action_schema(handler: _ToolFn) -> dict[str, Any]` returning `{"required": [...], "optional": {...}, "description": str}`
  - `_dispatch(group: str, actions: dict[str, _ToolFn], action: str, params: dict[str, Any] | None) -> dict[str, Any]` supporting `action == "help"`, `difflib.get_close_matches` hints on unknown actions and unknown params, expected-parameter hints on missing required params, and `ToolError` type-diagnostic hints when `params` values fail type checks.

- [ ] **Step 1: Write the failing tests in `tests/test_compound_tools.py`**

```python
async def test_compound_unknown_action_suggests_close_match():
    mcp = FastMCP("test")
    register_compound_tools(mcp)
    tool = await mcp.get_tool("word_document")
    with pytest.raises(ToolError, match="Did you mean 'create'"):
        await tool.fn(action="creat", params={})


async def test_compound_unknown_param_suggests_close_match_and_valid_params():
    mcp = FastMCP("test")
    register_compound_tools(mcp)
    tool = await mcp.get_tool("word_document")
    with pytest.raises(ToolError, match="Did you mean 'filename'"):
        await tool.fn(action="info", params={"filname": "doc.docx"})


async def test_compound_missing_required_param_lists_expected_signature():
    mcp = FastMCP("test")
    register_compound_tools(mcp)
    tool = await mcp.get_tool("word_content")
    with pytest.raises(ToolError, match="Required parameters:"):
        await tool.fn(action="add_heading", params={"filename": "doc.docx"})


async def test_compound_type_mismatch_raises_clear_tool_error():
    mcp = FastMCP("test")
    register_compound_tools(mcp)
    tool = await mcp.get_tool("word_content")
    with pytest.raises(ToolError, match="Invalid type for 'level'"):
        await tool.fn(
            action="add_heading",
            params={"filename": "doc.docx", "text": "Hi", "level": "one"},
        )


async def test_compound_help_action_returns_action_schemas():
    mcp = FastMCP("test")
    register_compound_tools(mcp)
    tool = await mcp.get_tool("word_document")
    overview = await tool.fn(action="help")
    assert overview["tool"] == "word_document"
    assert "create" in overview["actions"]

    detail = await tool.fn(action="help", params={"action": "create"})
    assert detail["action"] == "create"
    assert "filename" in detail["required"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_compound_tools.py -v`
Expected: FAIL on the 5 new tests (`Did you mean` / `Required parameters:` / `Invalid type` / `Unknown action 'help'`)

- [ ] **Step 3: Implement schema/type helpers and enhanced `_dispatch` in `src/word_mcp/tools/compound_tools.py`**

1. Add helper `_describe_action(handler: _ToolFn) -> dict[str, Any]` using `inspect.signature(handler)` and `typing.get_type_hints(handler)` (with `try/except Exception` fallback to raw `param.annotation`) to extract required parameter names/types, optional parameter names/types/defaults, and first line of docstring.
2. Add `_validate_param_types(group: str, action: str, handler: _ToolFn, kwargs: dict[str, Any]) -> None` using `pydantic.TypeAdapter` (cached per `(handler, param_name)`) or simple type checking so invalid parameter types (e.g., passing `"one"` for `level: int`) raise `ToolError(f"{group} action '{action}': Invalid type for '{name}': expected {expected}, got {type(val).__name__}")` before invoking `handler`. Note: use `TypeAdapter(annotation).validate_python(val, strict=False)` or check `ValidationError` so valid JSON-compatible values pass while genuinely incompatible types produce a clean `ToolError` with expected type info.
3. Update `_dispatch`:
   - If `action == "help"`: if `params` has `"action"`, return `_describe_action` for that action (or raise `ToolError` with `difflib.get_close_matches` if unknown); otherwise return `{"tool": group, "actions": {name: first_doc_line for name, fn in actions.items()}, "hint": "Call with action='help' and params={'action': '<name>'} for parameter details."}`.
   - On unknown `action`: include `difflib.get_close_matches(action, actions.keys(), n=2, cutoff=0.5)` suggestion (`Did you mean '...'?`) and mention `action='help'`.
   - On missing required param in `_Params.__missing__` or unknown param in `_Params.check_all_read`: include `difflib.get_close_matches` for unknown keys and list required/optional parameters from `_describe_action(handler)`.
4. Update `_PARAMS_NOTE` to mention `"help"` is available on every compound tool.

- [ ] **Step 4: Run tests and static checks to verify they pass**

Run: `uv run pytest tests/test_compound_tools.py -v`
Run: `uv run ruff check src tests scripts`
Run: `uv run mypy src`
Expected: PASS (all tests, ruff, and mypy pass)

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/tools/compound_tools.py tests/test_compound_tools.py
git commit -m "feat(compound): add help action, fuzzy suggestions, and parameter type hints"
```

---

### Task 4: Full Verification and Knowledge Graph Sync

**Files:**
- Verify: entire repository (`src/`, `tests/`, `scripts/`)

- [ ] **Step 1: Run full test suite, linter, and type checker**

Run:
```powershell
uv run pytest
uv run ruff check src tests scripts
uv run mypy src
```
Expected: All 133+ existing tests plus all new tests pass; `ruff` and `mypy` report zero issues.

- [ ] **Step 2: Update Graphify knowledge graph**

Run:
```powershell
uvx --from graphifyy graphify update .
```
Expected: `graphify-out/graph.json` and `graphify-out/GRAPH_REPORT.md` updated deterministically.

---

## Self-Review

1. **Spec coverage:** Bug fix (`OSError` subclass shadowing in `ExceptionTool`), improved error messages & `hint` guidance (`document_validators.py` + `compound_tools.py`), and accurate type/parameter hints (`_describe_action`, `_validate_param_types`, `action="help"` in `compound_tools.py`) are each covered by dedicated tasks and tests.
2. **Step scan:** Every step specifies exact file paths, function/method names, assertions, and verification commands.
3. **Type consistency:** `ExceptionTool.get_error_info`, `_validate_docx_path`, `_dispatch`, and `register_compound_tools` match the exact names in the codebase.
4. **Review Focus:** All five edge cases (`FileNotFoundError`/`PermissionError` ordering, `help` action with/without `params`, wrong types & misspelled action/param names, validator `"hint"` backwards compatibility, and compound tool count) are pinned by tests in Tasks 1-3.
5. **Proportion:** Concise task breakdown without duplicating full implementation code blocks.
