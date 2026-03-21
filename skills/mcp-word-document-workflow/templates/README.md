# Quick Templates

Prebuilt templates for fast Word document generation with this skill.

## How to Use

1. Pick a template by document type.
2. Replace uppercase placeholders in brackets.
3. Create the document with `word_create_document`.
4. Build content blocks with `word_add_heading` and `word_add_paragraph`.
5. Run QA before final delivery.

## Available Templates

- reporte-template.md
- carta-template.md
- acta-template.md
- propuesta-template.md
- informe-universitario-template.md
- solicitud-formal-template.md
- syllabus-template.md

## Common Reusable Fields

- [DOCUMENT_TITLE]
- [AUTHOR]
- [DATE]
- [AUDIENCE]
- [OBJECTIVE]
- [ABSOLUTE_OUTPUT_PATH]

## AI Prompt Starter

```text
Use [TEMPLATE_NAME] to generate a Word document.
Type: [DOCUMENT_TYPE]
Objective: [OBJECTIVE]
Audience: [AUDIENCE]
Tone: [LANGUAGE_AND_TONE]
Key inputs: [KEY_DATA]
Absolute output path: [ABSOLUTE_OUTPUT_PATH]
```

## Security Note

All output paths must be absolute and inside `MCP_ALLOWED_DIRECTORIES`.
