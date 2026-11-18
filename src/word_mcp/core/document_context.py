"""Document context management for Word Document Server.

This module provides atomic file persistence and a single-open context manager
for handling Word documents, ensuring crash-safe saves and consistent error mapping.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import tempfile
from collections.abc import Generator, Iterator
from contextlib import contextmanager
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.opc.exceptions import PackageNotFoundError

from word_mcp.exception import (
    DocumentProcessingError,
    FileOperationError,
    ValidationError,
)


@contextmanager
def atomic_write_target(target: str | Path) -> Iterator[Path]:
    """Stage writes to a same-directory temp file and replace atomically via os.replace.

    Creating the temp file in ``target_path.parent`` guarantees the rename stays on
    the same filesystem, and closing the OS descriptor before yielding avoids
    Windows file-locking errors during ``doc.save()`` or ``os.replace()``.
    """
    target_path = Path(target)
    parent = target_path.parent
    if not parent.exists():
        raise FileOperationError(
            f"Parent directory '{parent}' does not exist.",
            suggestion="Ensure the target directory exists before saving.",
        )

    fd, tmp_name = tempfile.mkstemp(
        dir=str(parent),
        prefix=f".{target_path.stem}.",
        suffix=target_path.suffix or ".tmp",
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        yield tmp_path
        if target_path.exists():
            with contextlib.suppress(OSError):
                shutil.copymode(target_path, tmp_path)
        try:
            os.replace(tmp_path, target_path)
        except PermissionError as exc:
            raise FileOperationError(
                f"Cannot replace target file '{target_path}': {exc}",
                suggestion="Close the document in Microsoft Word and verify write permissions.",
            ) from exc
    except BaseException:
        with contextlib.suppress(OSError):
            tmp_path.unlink(missing_ok=True)
        raise


@contextmanager
def document_context(
    filename: str | Path, mode: str = "read"
) -> Generator[DocumentType, None, None]:
    """Single-open Document lifecycle manager with atomic write on mode='write'."""
    target = os.fspath(filename)
    if mode not in ("read", "write"):
        raise ValidationError(f"Unsupported document_context mode: '{mode}'.")
    if not os.path.exists(target) and mode == "read":
        raise DocumentProcessingError(f"Document not found: {target}")

    try:
        if os.path.exists(target):
            doc = Document(target)
        else:
            doc = Document()

        yield doc

        if mode == "write":
            with atomic_write_target(target) as tmp_path:
                doc.save(str(tmp_path))

    except PackageNotFoundError as exc:
        raise DocumentProcessingError(f"Invalid Word document: {target}") from exc
    except (FileNotFoundError, PermissionError, FileOperationError):
        raise
    except OSError as error:
        raise FileOperationError(f"File system error: {error}") from error
