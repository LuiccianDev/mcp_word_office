# Document Editing Tools Reference

This reference describes the primary modification and inspection tools available in `office-mcp`.

## Inspection Tools

| Tool | Tokens | Purpose |
| :--- | :--- | :--- |
| `word_get_document_outline` | Ultra-Low (~100 tokens) | Extracts headings tree, level hierarchy, total paragraphs, and tables count. |
| `word_get_document_info` | Very Low (~150 tokens) | Metadata (author, title, modified date, word count estimate, sections). |
| `word_find_text_in_document` | Low (~200 tokens) | Searches a query substring and returns matching paragraph indexes and context snippets. |
| `word_get_document_text` | High (proportional to doc) | Dumps complete text. Use only when full semantic text analysis is strictly needed. |

## Mutation & Batch Tools

| Tool | Best Used For |
| :--- | :--- |
| `word_batch_add_content` | Adding multi-element sections (headings, paragraphs, tables, lists) in a single MCP request. |
| `word_search_and_replace` | Global or targeted string replacement across all paragraphs and tables. |
| `word_merge_documents` | Concatenating multiple `.docx` files with optional page breaks between them. |
| `word_add_hyperlink` | Inserting a web link with custom display text into a paragraph. |
| `word_add_bookmark` | Creating an internal navigation bookmark for cross-references. |
| `word_add_footnote` | Adding academic or legal citations linked to specific text. |
