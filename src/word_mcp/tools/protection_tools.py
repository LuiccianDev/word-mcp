"""Protection tools for Word Document Server.

These tools handle document protection features such as
password protection, restricted editing, and digital signatures.
"""

from __future__ import annotations

import datetime
from typing import Any

from word_mcp.core import (
    core_add_protection_info,
    core_create_signature_info,
    core_protect_document,
    core_unprotect_document,
    core_verify_document_protection,
    hash_password,
)
from word_mcp.core.document_context import document_context
from word_mcp.exception import ValidationError
from word_mcp.models.response_models import ProtectionResult
from word_mcp.validation.document_validators import (
    validate_docx_write,
    validate_file_read,
    validate_file_write,
)


@validate_docx_write("filename")
async def protect_document(filename: str, password: str) -> dict[str, Any]:
    """Add password protection to a Word document."""
    result = core_protect_document(filename, password)

    return ProtectionResult(
        status="success",
        filename=filename,
        protection_type="password",
        success=True,
        message=result["message"],
    ).model_dump()


@validate_file_write("filename")
async def unprotect_document(filename: str, password: str) -> dict[str, Any]:
    """Remove password protection from a Word document."""
    result = core_unprotect_document(filename, password)

    return ProtectionResult(
        status="success",
        filename=filename,
        protection_type="password",
        success=True,
        message=result["message"],
    ).model_dump()


@validate_docx_write("filename")
async def add_restricted_editing(
    filename: str, password: str, editable_sections: list[str]
) -> dict[str, Any]:
    """Add restricted editing to a Word document."""
    if not editable_sections:
        raise ValidationError(
            "No editable sections specified.",
            suggestion="Provide a non-empty list of editable section names.",
        )

    password_hash = hash_password(password)
    core_add_protection_info(
        filename,
        protection_type="restricted",
        password_hash=password_hash,
        sections=editable_sections,
    )

    return ProtectionResult(
        status="success",
        filename=filename,
        protection_type="restricted",
        success=True,
        message=f"Restricted editing applied. Editable: {', '.join(editable_sections)}",
    ).model_dump()


@validate_docx_write("filename")
async def add_digital_signature(
    filename: str, signer_name: str, reason: str | None = None
) -> dict[str, Any]:
    """Add a digital signature to a Word document."""
    with document_context(filename, mode="write") as doc:
        signature_info = core_create_signature_info(doc, signer_name, reason)

        doc.add_paragraph("")
        sig_para = doc.add_paragraph()
        sig_para.add_run(f"Digitally signed by: {signer_name}").bold = True
        if reason:
            sig_para.add_run(f"\nReason: {reason}")
        sig_para.add_run(
            f"\nDate: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        sig_para.add_run(f"\nSignature ID: {signature_info['content_hash'][:8]}")

    core_add_protection_info(
        filename,
        protection_type="signature",
        password_hash="",
        signature_info=signature_info,
    )

    return ProtectionResult(
        status="success",
        filename=filename,
        protection_type="signature",
        success=True,
        message="Digital signature added",
    ).model_dump()


@validate_file_read("filename")
async def verify_document(filename: str, password: str | None = None) -> dict[str, Any]:
    """Verify document protection and/or digital signature."""
    is_verified, message = core_verify_document_protection(filename, password)

    return ProtectionResult(
        status="success",
        filename=filename,
        protection_type="verification",
        success=is_verified,
        message=message,
    ).model_dump()
