---
name: office-document-editor
description: "Inspect, modify, search/replace, append batch content, or merge existing Word documents (.docx) efficiently using Office MCP. Prioritizes lightweight outline inspection and batch updates over chatty micro-calls."
argument-hint: "Target document path, edit action (inspect, search/replace, append batch content, merge), and modification parameters"
user-invocable: true
disable-model-invocation: false
---

# Office Document Editor

Use this skill to inspect and update **existing** Microsoft Word documents with surgical precision and minimal token consumption.

## Core Rules

1. **Inspect Before Reading Full Text:**
   - Use `word_get_document_outline` first to see headings, levels, paragraph counts, and tables.
   - Do NOT call `word_get_document_text` unless you strictly need to read the full body content.
2. **Batch Appending Over Single Injections:**
   - When adding new sections, appendices, or summary tables, use `word_batch_add_content`.
   - Never chain single `add_heading` + multiple `add_paragraph` + `add_table` calls.
3. **Surgical Search and Replace:**
   - Use `word_search_and_replace` for template variables (e.g. `{{CLIENT_NAME}}`, `[FECHA]`) or updating outdated terms across the entire document.
4. **Document Merging:**
   - When assembling multiple chapters or reports, use `word_merge_documents` rather than manually copying text between files.

## Workflow Patterns

### Pattern A: Inspect Structure
```python
# Call outline to get a compact summary (saves 95% tokens compared to get_document_text)
word_get_document_outline(filename="/path/to/documento.docx")
```

### Pattern B: Search & Replace
```python
# Replace placeholders or update terms
word_search_and_replace(
    filename="/path/to/contrato.docx",
    search_text="ACME Corp",
    replace_text="Global Logistics Inc."
)
```

### Pattern C: Append a Full New Section in 1 Call
```python
# Use batch_add_content to add a heading, paragraph, and table simultaneously
word_batch_add_content(
    filename="/path/to/informe.docx",
    elements=[
        {
            "type": "heading",
            "text": "Anexo A: Datos Complementarios",
            "level": 2
        },
        {
            "type": "paragraph",
            "text": "La siguiente tabla resume los indicadores complementarios auditados:"
        },
        {
            "type": "table",
            "headers": ["Métrica", "2025", "2026"],
            "rows": [
                ["Eficiencia", "84%", "92%"],
                ["Disponibilidad", "99.1%", "99.9%"]
            ]
        }
    ]
)
```

### Pattern D: Merge Documents
```python
# Merge appendices or chapters into the main document
word_merge_documents(
    primary_filename="/path/to/informe_final.docx",
    secondary_filenames=["/path/to/anexo1.docx", "/path/to/anexo2.docx"],
    add_page_break=True
)
```
