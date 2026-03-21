"""Protection tools for Word Document Server.

These tools handle document protection features such as
password protection, restricted editing, and digital signatures.
"""

import datetime
import hashlib
from typing import Any

from mcp_word.core import (
    core_add_protection_info,
    core_create_signature_info,
    core_verify_document_protection,
    core_protect_document,
    core_unprotect_document,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.models.response_models import ProtectionResult
from mcp_word.validation.document_validators import (
    validate_docx_read,
    validate_docx_write,
    validate_file_write,
)


@validate_docx_write("filename")
async def protect_document(filename: str, password: str) -> dict[str, Any]:
    """Add password protection to a Word document."""
    try:
        result = core_protect_document(filename, password)
        
        return ProtectionResult(
            status=result["status"],
            filename=filename,
            protection_type="password",
            success=result["status"] == "success",
            message=result["message"],
        ).model_dump()

    except Exception as error:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to protect document: {str(error)}"),
            filename=filename,
            operation="protect document",
        )


@validate_file_write("filename")
async def unprotect_document(filename: str, password: str) -> dict[str, Any]:
    """Remove password protection from a Word document."""
    try:
        result = core_unprotect_document(filename, password)
        
        return ProtectionResult(
            status=result["status"],
            filename=filename,
            protection_type="password",
            success=result["status"] == "success",
            message=result["message"],
        ).model_dump()

    except Exception as error:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to unprotect document: {str(error)}"),
            filename=filename,
            operation="unprotect document",
        )


@validate_docx_write("filename")
async def add_restricted_editing(
    filename: str, password: str, editable_sections: list[str]
) -> dict[str, Any]:
    """Add restricted editing to a Word document."""
    try:
        if not editable_sections:
            return {
                "status": "error",
                "message": "No editable sections specified.",
            }

        password_hash = hashlib.sha256(password.encode()).hexdigest()
        success = core_add_protection_info(
            filename,
            protection_type="restricted",
            password_hash=password_hash,
            sections=editable_sections,
        )

        return ProtectionResult(
            status="success" if success else "error",
            filename=filename,
            protection_type="restricted",
            success=success,
            message=f"Restricted editing {'applied' if success else 'failed'}. Editable: {', '.join(editable_sections)}",
        ).model_dump()

    except Exception as error:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to add restricted editing: {str(error)}"),
            filename=filename,
            operation="add restricted editing",
        )


@validate_docx_write("filename")
async def add_digital_signature(
    filename: str, signer_name: str, reason: str | None = None
) -> dict[str, Any]:
    """Add a digital signature to a Word document."""
    try:
        with document_context(filename, mode="write") as doc:
            signature_info = core_create_signature_info(doc, signer_name, reason)
            
            # Visible signature block
            doc.add_paragraph("")
            sig_para = doc.add_paragraph()
            sig_para.add_run(f"Digitally signed by: {signer_name}").bold = True
            if reason:
                sig_para.add_run(f"\nReason: {reason}")
            sig_para.add_run(f"\nDate: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            sig_para.add_run(f"\nSignature ID: {signature_info['content_hash'][:8]}")

        success = core_add_protection_info(
            filename,
            protection_type="signature",
            password_hash="",
            signature_info=signature_info,
        )

        return ProtectionResult(
            status="success" if success else "error",
            filename=filename,
            protection_type="signature",
            success=success,
            message=f"Digital signature {'added' if success else 'failed'}",
        ).model_dump()

    except Exception as error:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to add digital signature: {str(error)}"),
            filename=filename,
            operation="add digital signature",
        )


@validate_docx_read("filename")
async def verify_document(filename: str, password: str | None = None) -> dict[str, Any]:
    """Verify document protection and/or digital signature."""
    try:
        is_verified, message = core_verify_document_protection(filename, password)
        
        return ProtectionResult(
            status="success" if is_verified else "error",
            filename=filename,
            protection_type="verification",
            success=is_verified,
            message=message,
        ).model_dump()

    except Exception as error:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to verify document: {str(error)}"),
            filename=filename,
            operation="verify document",
        )
