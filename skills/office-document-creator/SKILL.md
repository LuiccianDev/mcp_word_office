---
name: office-document-creator
description: "Create high-quality, professional Word documents (.docx) in a single pass using Office MCP. Generates complete documents with typography, margins, line spacing, header/footer, page numbers, headings, paragraphs, and tables in 1 tool call, saving 90% tokens."
argument-hint: "Document type, objective, design preset ('modern', 'formal', 'academic'), and target path"
user-invocable: true
disable-model-invocation: false
---

# Office Document Creator

Use this skill to create complete, visually styled Microsoft Word documents in **1 single tool call**, avoiding slow micro-iterations and preserving model context.

## Core Rules

1. **Always use Single-Pass Construction for new documents:**
   - **`word_build_document`**: Primary tool. Use when creating structured reports, proposals, letters, minutes, or academic papers that include headers, tables, bullet lists, or specific margins.
   - **`word_create_from_markdown`**: Fast-path tool. Use when the user requests direct drafting of structured markdown text (headers `#`, lists `-`, tables `|`, bold `**`).
2. **Never make micro-calls for new documents:**
   - Do NOT chain `create_document` -> `add_heading` -> `add_paragraph` -> `add_table`.
   - Send the entire document structure inside the `elements` array or `markdown_content` in **one call**.
3. **Always use absolute allowed paths** within `MCP_ALLOWED_DIRECTORIES`.

## Design Presets

| Preset | Font | Size | Line Spacing | Margins | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `modern` (Default) | Calibri | 11.0 pt | 1.15 | 2.5 cm | Executive summaries, business proposals, corporate reports |
| `formal` | Times New Roman | 12.0 pt | 1.50 | Top/Bottom 2.5cm, Left 3.0cm | Legal contracts, official letters, meeting minutes |
| `academic` | Times New Roman | 12.0 pt | 2.00 (Double) | 2.54 cm (1 inch) | University reports, theses, research papers |
| `compact` | Arial | 10.0 pt | 1.00 (Single) | 2.0 cm | Invoices, datasheets, quick memos |

## Canonical Usage Pattern: `word_build_document`

```json
{
  "filename": "/path/to/informe_ventas.docx",
  "title": "Informe Trimestral de Ventas",
  "author": "Equipo Comercial",
  "preset": "modern",
  "header_text": "Informe Q1 2026 | Confidencial",
  "footer_text": "Empresa S.A.",
  "page_numbers": true,
  "elements": [
    {
      "type": "heading",
      "text": "1. Resumen Ejecutivo",
      "level": 1
    },
    {
      "type": "paragraph",
      "text": "Durante el primer trimestre se alcanzó un crecimiento del **18%** respecto al año anterior.",
      "alignment": "justify",
      "space_after": 6
    },
    {
      "type": "heading",
      "text": "2. Resultados Clave",
      "level": 2
    },
    {
      "type": "table",
      "headers": ["Región", "Meta", "Alcanzado", "Cumplimiento"],
      "rows": [
        ["Norte", "$100k", "$112k", "112%"],
        ["Sur", "$80k", "$85k", "106%"],
        ["Centro", "$120k", "$128k", "107%"]
      ],
      "style": "Table Grid"
    },
    {
      "type": "page_break"
    },
    {
      "type": "heading",
      "text": "3. Conclusiones y Próximos Pasos",
      "level": 2
    },
    {
      "type": "bullet_list",
      "items": [
        "Consolidar la expansión regional en el Q2.",
        "Aumentar la inversión en automatización de procesos."
      ]
    }
  ]
}
```

## Quick Verification (QA)

After generating the document with `word_build_document`:
- Run `word_get_document_outline(filename)` to confirm the structure without reloading the full text.
- Present a concise, helpful confirmation to the user.
