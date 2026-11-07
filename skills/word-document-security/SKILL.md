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
   - Run `word_verify_document` before applying cryptographic locks to ensure the XML structure is valid and uncorrupted.
2. **Password Protection Modes:**
   - **`read_only`**: Prevents any changes without providing the password.
   - **`comments`**: Allows reviewers to insert comments only; body text cannot be altered.
   - **`tracked_changes`**: Forces all revisions to be tracked with author metadata.
3. **Distribution & PDF Conversion:**
   - For finalized documents intended for signing or non-editable distribution, export via `word_convert_to_pdf`.

## Usage Patterns

### Pattern A: Lock Document with Password
```python
# Protect document against unauthorized alterations
word_protect_document(
    filename="/path/to/contrato_final.docx",
    password="SecurePassword2026!",
    protection_type="read_only"
)
```

### Pattern B: Unlock Protected Document
```python
# Remove editing lock using the correct password
word_unprotect_document(
    filename="/path/to/contrato_final.docx",
    password="SecurePassword2026!"
)
```

### Pattern C: Restrict Editing to Comments Only
```python
# Enable review mode where users can only leave comments
word_add_restricted_editing(
    filename="/path/to/borrador_revisores.docx",
    editing_mode="comments"
)
```

### Pattern D: Export to PDF for Distribution
```python
# Convert completed .docx to PDF
word_convert_to_pdf(
    filename="/path/to/informe_anual.docx",
    output_pdf_path="/path/to/informe_anual.pdf"
)
```
