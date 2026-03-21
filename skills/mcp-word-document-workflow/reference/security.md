# Security: Strict Allowed-Path Policy

Use this guide to enforce safe file operations for all Word workflows.

## Core Rule

- Only use absolute paths inside `MCP_ALLOWED_DIRECTORIES`.
- If a path is invalid or out of scope: stop and request a valid path.
- Never fallback to relative paths.

## Validate Before Write Operations

Always validate path safety before:
- `word_create_document`
- `word_copy_document` (destination)
- `word_merge_documents` (destination)
- `word_convert_to_pdf` (output)

## Validation Checklist

- [ ] Path is absolute
- [ ] Path is inside `MCP_ALLOWED_DIRECTORIES`
- [ ] Parent directory exists (or user approved creation)
- [ ] Write permission exists
- [ ] File extension matches operation (`.docx` or `.pdf`)

## Failure Behavior

If validation fails:
1. Do not create or modify any files.
2. Return a clear error explaining why the path is invalid.
3. Ask for a valid absolute path.

## Error Message Pattern

```text
Path validation failed: [PATH]
Reason: path is outside MCP_ALLOWED_DIRECTORIES.
Please provide an absolute path inside an allowed directory.
```

## Operational Guidance

- Validate early, before building content.
- Keep user informed about final output paths.
- Prefer explicit and actionable error messages.
