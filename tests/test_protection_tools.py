"""Tests for document protection and PBKDF2 password hashing."""

import pytest
from docx import Document

from word_mcp.core.protection import (
    core_add_protection_info,
    core_verify_document_protection,
    hash_password,
    verify_password,
)
from word_mcp.exception.unprotect import remove_protection_info
from word_mcp.tools.protection_tools import add_restricted_editing


class TestPasswordHashing:
    """Test PBKDF2 password hashing and verification."""

    def test_hash_password_format_and_salt_uniqueness(self):
        pwd = "SecretPassword123!"
        hash1 = hash_password(pwd)
        hash2 = hash_password(pwd)

        assert hash1.startswith("pbkdf2_sha256$")
        assert hash2.startswith("pbkdf2_sha256$")
        # Different salts must yield different hashes
        assert hash1 != hash2

    def test_verify_password_success(self):
        pwd = "MyStrongPassword#2026"
        hashed = hash_password(pwd)
        assert verify_password(pwd, hashed) is True

    def test_verify_password_failure(self):
        pwd = "MyStrongPassword#2026"
        hashed = hash_password(pwd)
        assert verify_password("WrongPassword", hashed) is False

    def test_verify_password_edge_cases(self):
        assert verify_password("", "pbkdf2_sha256$100000$salt$hash") is False
        assert verify_password("pwd", "") is False
        assert verify_password("pwd", "invalid_format") is False
        assert verify_password("pwd", "pbkdf2_sha256$invalid_int$salt$hash") is False
        assert verify_password("pwd", "sha256$somehash") is False


class TestRestrictedEditingProtection:
    """Test restricted editing workflow with secure password hashing."""

    @pytest.mark.asyncio
    async def test_restricted_editing_flow(self, tmp_path):
        doc_file = tmp_path / "test_doc.docx"
        doc = Document()
        doc.add_paragraph("Section 1 content")
        doc.add_paragraph("Section 2 content")
        doc.save(str(doc_file))

        # Add restricted editing with password
        res = await add_restricted_editing(
            filename=str(doc_file),
            password="RestrictedPassword456",
            editable_sections=["Section 1"],
        )
        assert res["status"] == "success"
        assert res["success"] is True

        # Verify protection with correct password
        is_valid, msg = core_verify_document_protection(
            str(doc_file), password="RestrictedPassword456"
        )
        assert is_valid is True
        assert "restricted" in msg

        # Verify protection with incorrect password
        is_valid, msg = core_verify_document_protection(
            str(doc_file), password="IncorrectPassword"
        )
        assert is_valid is False
        assert msg == "Incorrect password"

    def test_remove_protection_info_flow(self, tmp_path):
        doc_file = tmp_path / "test_unprotect.docx"
        doc = Document()
        doc.add_paragraph("Sample content")
        doc.save(str(doc_file))

        pwd = "UnprotectKey789"
        hashed = hash_password(pwd)
        core_add_protection_info(
            str(doc_file),
            protection_type="restricted",
            password_hash=hashed,
            sections=["Intro"],
        )

        metadata_file = tmp_path / "test_unprotect.protection"
        assert metadata_file.exists()

        # Try removing with wrong password
        success, msg = remove_protection_info(str(doc_file), password="Wrong")
        assert success is False
        assert msg == "Incorrect password"
        assert metadata_file.exists()

        # Remove with correct password
        success, msg = remove_protection_info(str(doc_file), password=pwd)
        assert success is True
        assert "successfully" in msg
        assert not metadata_file.exists()
