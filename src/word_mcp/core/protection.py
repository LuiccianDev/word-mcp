"""
Document protection functionality for Word Document Server.
"""

import datetime
import hashlib
import hmac
import io
import json
import logging
import os
import secrets
from pathlib import Path
from typing import Any

from docx.document import Document

from word_mcp.core.document_context import atomic_write_target, document_context
from word_mcp.exception import (
    DocumentProcessingError,
    FileOperationError,
    SecurityError,
    ValidationError,
)
from word_mcp.utils.security import resolve_sandboxed_path


logger = logging.getLogger(__name__)

PBKDF2_ALGORITHM = "sha256"
PBKDF2_ITERATIONS = 100_000
MIN_PBKDF2_ITERATIONS = 100_000
MAX_PBKDF2_ITERATIONS = 2_000_000
MAX_METADATA_BYTES = 64 * 1024


def hash_password(password: str, salt: str | None = None) -> str:
    """
    Hash a password securely using PBKDF2-HMAC-SHA256 with a cryptographic salt.

    Args:
        password: The plain text password to hash
        salt: Optional hexadecimal salt string (generated randomly if omitted)

    Returns:
        Formatted string: pbkdf2_sha256$<iterations>$<salt_hex>$<hash_hex>
    """
    if not salt:
        salt = secrets.token_hex(16)

    dk = hashlib.pbkdf2_hmac(
        PBKDF2_ALGORITHM,
        password.encode("utf-8"),
        salt.encode("utf-8"),
        PBKDF2_ITERATIONS,
    )
    return f"pbkdf2_sha256${PBKDF2_ITERATIONS}${salt}${dk.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    """
    Verify a password against a stored PBKDF2 hash using constant-time comparison.

    Args:
        password: Plain text password to verify
        stored_hash: Stored hash formatted as pbkdf2_sha256$<iterations>$<salt>$<hash>

    Returns:
        True if password matches, False otherwise
    """
    if not stored_hash or not password:
        return False

    if stored_hash.startswith("pbkdf2_sha256$"):
        parts = stored_hash.split("$")
        if len(parts) == 4:
            _, iter_str, salt, expected_hash = parts
            try:
                iterations = int(iter_str)
                if not (MIN_PBKDF2_ITERATIONS <= iterations <= MAX_PBKDF2_ITERATIONS):
                    return False
                dk = hashlib.pbkdf2_hmac(
                    PBKDF2_ALGORITHM,
                    password.encode("utf-8"),
                    salt.encode("utf-8"),
                    iterations,
                )
                return hmac.compare_digest(dk.hex(), expected_hash)
            except (ValueError, TypeError):
                return False

    return False


def _resolve_metadata_path(doc_path: str) -> Path:
    """Resolve sibling .protection sidecar path within WORD_MCP_DIRS, rejecting symlinks."""
    base_path, _ = os.path.splitext(doc_path)
    raw_meta = f"{base_path}.protection"
    raw_path_obj = Path(raw_meta)
    resolved_parent = resolve_sandboxed_path(
        raw_path_obj.parent, param_name="metadata_path"
    )
    unresolved_leaf = resolved_parent / raw_path_obj.name
    if unresolved_leaf.is_symlink():
        raise SecurityError(
            f"Protection metadata path '{raw_meta}' must not be a symlink.",
            suggestion="Remove the symlinked .protection file and re-apply protection.",
        )
    return resolve_sandboxed_path(unresolved_leaf, param_name="metadata_path")


def _read_protection_metadata(metadata_path: Path) -> dict[str, Any]:
    """Read and size-bound a .protection JSON sidecar safely."""
    if metadata_path.is_symlink():
        raise SecurityError(
            f"Protection metadata path '{metadata_path}' must not be a symlink.",
            suggestion="Remove the symlinked .protection file and re-apply protection.",
        )
    if metadata_path.stat().st_size > MAX_METADATA_BYTES:
        raise DocumentProcessingError(
            f"Protection metadata '{metadata_path}' exceeds maximum allowed size."
        )
    with open(metadata_path, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise DocumentProcessingError("Protection metadata must be a JSON object.")
    return data


def _compute_document_content_hash(doc: Document) -> str:
    """Hash body paragraphs, table cells, and headers/footers, excluding trailing signature block."""
    para_texts = [p.text for p in doc.paragraphs]
    if (
        para_texts
        and "Digitally signed by: " in para_texts[-1]
        and "Signature ID: " in para_texts[-1]
    ):
        para_texts.pop()
        if para_texts and para_texts[-1] == "":
            para_texts.pop()
    table_texts = [
        cell.text for table in doc.tables for row in table.rows for cell in row.cells
    ]
    header_footer_texts = [
        p.text
        for section in doc.sections
        for hf in (section.header, section.footer)
        for p in hf.paragraphs
    ]
    combined = "\n".join(para_texts + table_texts + header_footer_texts)
    return hashlib.sha256(combined.encode("utf-8")).hexdigest()


def _encrypt_docx_file(
    doc_path: str, password: str, *, reject_if_encrypted: bool = False
) -> None:
    """Encrypt a .docx file in-place atomically via msoffcrypto."""
    import msoffcrypto

    try:
        with open(doc_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            if reject_if_encrypted and office_file.is_encrypted():
                raise DocumentProcessingError(
                    "Document is already encrypted.",
                    suggestion="Unprotect the document before applying new password encryption.",
                )
            encrypted_io = io.BytesIO()
            office_file.encrypt(password=password, outfile=encrypted_io)

        with atomic_write_target(doc_path) as tmp_path:
            tmp_path.write_bytes(encrypted_io.getvalue())
    except (
        SecurityError,
        DocumentProcessingError,
        FileOperationError,
        ValidationError,
    ):
        raise
    except OSError as exc:
        raise FileOperationError(f"Failed to encrypt document: {exc}") from exc
    except Exception as exc:
        raise DocumentProcessingError(f"Failed to encrypt document: {exc}") from exc


def core_add_protection_info(
    doc_path: str,
    protection_type: str,
    password_hash: str,
    sections: list[str] | None = None,
    signature_info: dict[str, Any] | None = None,
    raw_password: str | None = None,
) -> bool:
    """
    Add document protection information to a separate metadata file and encrypt the document.

    Args:
        doc_path: Path to the document
        protection_type: Type of protection ('password', 'restricted', 'signature')
        password_hash: Hashed password for security
        sections: List of section names that can be edited (for restricted editing)
        signature_info: Information about digital signature
        raw_password: The actual password for document encryption

    Returns:
        True if protection info was successfully added, False otherwise
    """
    metadata_path = _resolve_metadata_path(doc_path)

    # Prepare protection data
    protection_data: dict[str, Any] = {
        "type": protection_type,
        "password_hash": password_hash,
        "applied_date": datetime.datetime.now().isoformat(),
    }

    if sections:
        protection_data["editable_sections"] = sections

    if signature_info:
        protection_data["signature"] = signature_info

    # Encrypt before writing metadata so .protection is persisted in a single atomic pass.
    if protection_type == "password" and raw_password:
        _encrypt_docx_file(doc_path, raw_password)
        protection_data["true_encryption"] = True

    with atomic_write_target(metadata_path) as meta_tmp:
        with open(meta_tmp, "w", encoding="utf-8") as f:
            json.dump(protection_data, f, indent=2)

    return True


def core_protect_document(filename: str, password: str) -> dict[str, Any]:
    """Add password protection to a Word document using msoffcrypto."""
    metadata_path = _resolve_metadata_path(filename)
    _encrypt_docx_file(filename, password, reject_if_encrypted=True)

    try:
        if metadata_path.exists() or metadata_path.is_symlink():
            os.remove(metadata_path)
    except OSError as e:
        raise FileOperationError(
            f"Failed to remove stale protection metadata: {e}"
        ) from e

    return {"status": "success", "message": "Document encrypted successfully."}


def core_unprotect_document(filename: str, password: str) -> dict[str, Any]:
    """Remove password protection from a Word document."""
    import msoffcrypto

    try:
        with open(filename, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            if not office_file.is_encrypted():
                raise DocumentProcessingError(
                    "Document is not encrypted.",
                    suggestion="No password decryption is needed for this document.",
                )
            office_file.load_key(password=password)
            decrypted_io = io.BytesIO()
            office_file.decrypt(outfile=decrypted_io)

        with atomic_write_target(filename) as tmp_path:
            tmp_path.write_bytes(decrypted_io.getvalue())

        return {"status": "success", "message": "Document decrypted successfully."}
    except msoffcrypto.exceptions.InvalidKeyError as exc:
        raise ValidationError(
            "Incorrect password.",
            suggestion="Verify the document password and try again.",
        ) from exc
    except (DocumentProcessingError, FileOperationError, ValidationError):
        raise
    except OSError as e:
        raise FileOperationError(f"Failed to decrypt document: {e}") from e
    except Exception as e:
        raise DocumentProcessingError(f"Failed to decrypt document: {e}") from e


def core_verify_document_protection(
    doc_path: str, password: str | None = None
) -> tuple[bool, str]:
    """
    Verify if a document is protected and if the password is correct.

    Args:
        doc_path: Path to the document
        password: Password to verify

    Returns:
        Tuple of (is_protected_and_verified, message)
    """
    metadata_path = _resolve_metadata_path(doc_path)

    # Check if protection metadata exists
    if not metadata_path.exists():
        import msoffcrypto

        try:
            with open(doc_path, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                if office_file.is_encrypted():
                    if password is not None:
                        if not password:
                            return False, "Incorrect password"
                        try:
                            office_file.load_key(password=password)
                            office_file.decrypt(outfile=io.BytesIO())
                            return (
                                True,
                                "Document is password-encrypted and password verified",
                            )
                        except msoffcrypto.exceptions.InvalidKeyError:
                            return False, "Incorrect password"
                    return True, "Document is protected with password protection"
        except msoffcrypto.exceptions.FileFormatError as exc:
            raise DocumentProcessingError(
                f"File '{doc_path}' is not a valid Word document (.docx)."
            ) from exc
        except OSError as exc:
            raise FileOperationError(
                f"Cannot read document '{doc_path}': {exc}"
            ) from exc
        return False, "Document is not protected"

    try:
        protection_data = _read_protection_metadata(metadata_path)
        protection_type = str(protection_data.get("type", "unknown"))

        if password is not None and not (
            protection_type == "signature" and not protection_data.get("password_hash")
        ):
            stored_hash = protection_data.get("password_hash", "")
            if not password or not verify_password(password, stored_hash):
                return False, "Incorrect password"

        if protection_type == "signature":
            return core_verify_signature(doc_path)

        return True, f"Document is protected with {protection_type} protection"

    except SecurityError:
        raise
    except (OSError, json.JSONDecodeError, DocumentProcessingError) as e:
        return False, f"Error verifying protection: {str(e)}"


def core_is_section_editable(doc_path: str, section_name: str) -> bool:
    """
    Check if a specific section of a document is editable.

    Args:
        doc_path: Path to the document
        section_name: Name of the section to check

    Returns:
        True if section is editable, False otherwise
    """
    metadata_path = _resolve_metadata_path(doc_path)

    if not metadata_path.exists():
        return True

    try:
        protection_data = _read_protection_metadata(metadata_path)

        if protection_data.get("type") != "restricted":
            return protection_data.get("type") != "password"

        editable_sections: list[str] = protection_data.get("editable_sections", [])
        return section_name in editable_sections

    except SecurityError:
        raise
    except (OSError, json.JSONDecodeError, KeyError, DocumentProcessingError):
        return False


def core_create_signature_info(
    doc: Document, signer_name: str, reason: str | None = None
) -> dict[str, Any]:
    """
    Create signature information for a document.

    Args:
        doc: Document object
        signer_name: Name of the person signing the document
        reason: Optional reason for signing

    Returns:
        Dictionary containing signature information
    """
    signature_info: dict[str, Any] = {
        "signer": signer_name,
        "timestamp": datetime.datetime.now().isoformat(),
    }

    if reason:
        signature_info["reason"] = reason

    signature_info["content_hash"] = _compute_document_content_hash(doc)
    return signature_info


def core_verify_signature(doc_path: str) -> tuple[bool, str]:
    """
    Verify a document's digital signature.

    Args:
        doc_path: Path to the document

    Returns:
        Tuple of (is_valid, message)
    """
    metadata_path = _resolve_metadata_path(doc_path)

    if not metadata_path.exists():
        return False, "Document is not signed"

    try:
        protection_data = _read_protection_metadata(metadata_path)

        if protection_data.get("type") != "signature":
            return (
                False,
                f"Document is protected with {protection_data.get('type')} protection, not a signature",
            )

        signature_info = protection_data.get("signature", {})
        original_hash = signature_info.get("content_hash")

        if not original_hash:
            return False, "Invalid signature: missing content hash"

        with document_context(doc_path, mode="read") as doc:
            current_hash = _compute_document_content_hash(doc)
            last_para = doc.paragraphs[-1].text if doc.paragraphs else ""

        expected_signer = str(signature_info.get("signer", ""))
        if current_hash != original_hash or (
            expected_signer
            and (
                f"Digitally signed by: {expected_signer}" not in last_para
                or f"Signature ID: {original_hash[:8]}" not in last_para
            )
        ):
            return (
                False,
                f"Document has been modified since it was signed by {signature_info.get('signer')}",
            )

        return (
            True,
            f"Document signature is valid. Signed by {signature_info.get('signer')} on {signature_info.get('timestamp')}",
        )

    except SecurityError:
        raise
    except (
        OSError,
        json.JSONDecodeError,
        KeyError,
        DocumentProcessingError,
        FileOperationError,
    ) as e:
        return False, f"Error verifying signature: {str(e)}"


def remove_protection_info(
    filename: str, password: str | None = None
) -> tuple[bool, str]:
    """Remove protection metadata and decrypt the document if encrypted."""
    try:
        metadata_path = _resolve_metadata_path(filename)
    except (SecurityError, ValidationError) as exc:
        return False, f"Error removing protection: {exc}"

    if not metadata_path.exists():
        return False, "Document is not protected"

    try:
        protection_data = _read_protection_metadata(metadata_path)

        if protection_data.get("password_hash"):
            if not password or not verify_password(
                password, protection_data.get("password_hash", "")
            ):
                return False, "Incorrect password"

        if protection_data.get("true_encryption"):
            if not password:
                return False, "Incorrect password"
            try:
                core_unprotect_document(filename, password)
            except (
                ValidationError,
                DocumentProcessingError,
                FileOperationError,
            ) as exc:
                return False, f"Failed to decrypt document: {exc}"

        os.remove(metadata_path)
        return True, "Protection removed successfully"
    except Exception as exc:
        return False, f"Error removing protection: {exc}"
