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
   - Run verify_document to inspect sidecar protection metadata before distribution.
2. **Protection Modes:**
   - **Container encryption**: `protect_document` encrypts the `.docx` with a password.
   - **Restricted editing metadata**: `add_restricted_editing` records an advisory list of editable sections in the `.protection` sidecar. Microsoft Word does not enforce this list.
   - **Signature marker**: `add_digital_signature` writes a visible signature block and a SHA-256 content hash. It is not a cryptographic signature.
3. **Distribution & PDF Conversion:**
   - For finalized documents intended for signing or non-editable distribution, export via `convert_to_pdf`.

## Usage Patterns

### Pattern A: Lock Document with Password
```python
# Protect document against unauthorized alterations
protect_document(
    filename="/path/to/contrato_final.docx",
    password="SecurePassword2026!"
)
```

### Pattern B: Unlock Protected Document
```python
# Remove editing lock using the correct password
unprotect_document(
    filename="/path/to/contrato_final.docx",
    password="SecurePassword2026!"
)
```

### Pattern C: Restrict Editing Sections
```python
add_restricted_editing(
    filename="/path/to/borrador_revisores.docx",
    password="SecurePassword2026!",
    editable_sections=["Introduccion", "Conclusiones"]
)
```

### Pattern D: Export to PDF for Distribution
```python
# Convert completed .docx to PDF
convert_to_pdf(
    filename="/path/to/informe_anual.docx",
    output_filename="/path/to/informe_anual.pdf"
)
```
