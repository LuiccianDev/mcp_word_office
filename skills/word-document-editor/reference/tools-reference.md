# Document Editing Tools Reference

This reference describes the primary modification and inspection tools available in `word-mcp`.

## Inspection Tools

| Tool | Tokens | Purpose |
| :--- | :--- | :--- |
| `get_document_outline` | Ultra-Low (~100 tokens) | Extracts headings tree, level hierarchy, total paragraphs, and tables count. |
| `get_document_info` | Very Low (~150 tokens) | Metadata (author, title, modified date, word count estimate, sections). |
| `find_text_in_document` | Low (~200 tokens) | Searches a query substring and returns matching paragraph indexes and context snippets. |
| `get_document_text` | High (proportional to doc) | Dumps complete text. Use only when full semantic text analysis is strictly needed. |

## Mutation & Batch Tools

| Tool | Best Used For |
| :--- | :--- |
| `batch_add_content` | Adding multi-element sections (headings, paragraphs, tables, lists) in a single MCP request. |
| `search_and_replace` | Global or targeted string replacement across all paragraphs and tables. |
| `merge_documents` | Concatenating multiple `.docx` files with optional page breaks between them. |
| `add_hyperlink` | Inserting a web link with custom display text into a paragraph. |
| `add_bookmark` | Creating an internal navigation bookmark for cross-references. |
| `add_footnote_to_document` | Adding academic or legal citations linked to specific text. |
