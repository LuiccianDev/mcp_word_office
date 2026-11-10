# Word MCP - Tools Reference

This document provides a comprehensive reference for all tools available in the Word MCP server. These tools allow you to create, modify, and manage Word documents programmatically.

> **Note:** Tool names are descriptive and do not use a namespace prefix (for example, `create_document`).

## Table of Contents

- [Single-Pass & Batch Tools (Recommended)](#single-pass--batch-tools-recommended)
- [Document Tools](#document-tools)
- [Content Tools](#content-tools)
- [Format Tools](#format-tools)
- [Protection Tools](#protection-tools)
- [Footnote Tools](#footnote-tools)
- [Extended Document Inspection](#extended-document-inspection)
- [Header & Footer](#header--footer)
- [Hyperlinks & Bookmarks](#hyperlinks--bookmarks)
- [Property & Layout Tools](#property--layout-tools)

---

## Single-Pass & Batch Tools (Recommended)

These high-performance tools generate or modify documents in **1 single MCP call**, saving over 90% of model tokens and conversational turns compared to legacy micro-calls.

### `build_document`

```python
build_document(filename: str | None = None, file_path: str | None = None, title: str | None = None, author: str | None = None, preset: str | None = None, font_name: str | None = None, font_size: float | None = None, line_spacing: float | None = None, space_after: float | None = None, margins: dict[str, float] | None = None, orientation: str | None = None, header_text: str | None = None, footer_text: str | None = None, page_numbers: bool = False, elements: list[dict[str, typing.Any]] | None = None) -> dict[str, typing.Any]
```

Build a complete, professionally styled Word document in a single tool call with typography, margins, headers, footers, page numbers, headings, paragraphs, and tables.

**Parameters:**

- `filename` (str): Output document file path (accepts `filename` or `file_path`).
- `title` (str, optional): Document title metadata and core properties.
- `author` (str, optional): Document author metadata.
- `preset` (str, optional): Visual style preset (`"modern"`, `"formal"`, `"academic"`, `"compact"`). Default: `"modern"`.
- `margins` (dict, optional): Custom margins `{"top": float, "bottom": float, "left": float, "right": float}` in centimeters.
- `orientation` (str, optional): Page orientation (`"portrait"` or `"landscape"`). Default: `"portrait"`.
- `font_name` (str, optional): Base body font family (e.g. `"Calibri"`, `"Times New Roman"`, `"Arial"`).
- `font_size` (float, optional): Base body font size in points (e.g. 11.0, 12.0).
- `line_spacing` (float, optional): Line spacing multiplier (e.g. 1.15, 1.5, 2.0).
- `space_after` (float, optional): Spacing after paragraphs in points (e.g. 6.0).
- `header_text` (str, optional): Running header text for document pages.
- `footer_text` (str, optional): Running footer text for document pages.
- `page_numbers` (bool, optional): If `True`, injects automatic Word dynamic page numbering (`PAGE`) in the footer. Default: `False`.
- `elements` (list of dict): Ordered list of document elements to render:
  - `{"type": "heading", "text": "...", "level": 1-9}`
  - `{"type": "paragraph", "text": "...", "alignment": "left|center|right|justify", "space_after": 6}` (alignment Default: `left`, supports `**bold**`, `*italic*`)
  - `{"type": "table", "headers": ["Col1", "Col2"], "rows": [["A", "B"]], "style": "Table Grid"}`
  - `{"type": "bullet_list", "items": ["Item 1", "Item 2"]}`
  - `{"type": "numbered_list", "items": ["Step 1", "Step 2"]}`
  - `{"type": "callout", "text": "...", "title": "NOTE"}` (callout `title` is accepted but not rendered)
  - `{"type": "page_break"}`

**Returns:**
- `BuildDocumentResult`: Structured summary including `status`, `filename`, `total_elements`, and elements breakdown.

---

### `create_from_markdown`

```python
create_from_markdown(markdown_content: str, filename: str | None = None, file_path: str | None = None, title: str | None = None, author: str | None = None, preset: str | None = None, font_name: str | None = None, font_size: float | None = None, line_spacing: float | None = None, space_after: float | None = None, margins: dict[str, float] | None = None, header_text: str | None = None, footer_text: str | None = None, page_numbers: bool = False) -> dict[str, typing.Any]
```

Quickly draft and format a Word document from structured markdown text in a single pass.

**Parameters:**

- `markdown_content` (str): Markdown text containing headings (`#`, `##`), bullet lists (`-`), numbered lists (`1.`), tables (`| Col1 | Col2 |`), callouts (`>`), and inline formatting (`**bold**`, `*italic*`).
- `filename` (str, optional): Output document file path (accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `title` (str, optional): Document title metadata.
- `author` (str, optional): Document author metadata.
- `preset` (str, optional): Visual style preset (`"modern"`, `"formal"`, `"academic"`, `"compact"`).
- `font_name` (str, optional): Base body font family.
- `font_size` (float, optional): Base body font size in points.
- `line_spacing` (float, optional): Line spacing multiplier.
- `space_after` (float, optional): Space after paragraphs in points.
- `margins` (dict, optional): Custom margins `{"top": float, "bottom": float, "left": float, "right": float}` in centimeters.
- `header_text` (str, optional): Running header text.
- `footer_text` (str, optional): Running footer text.
- `page_numbers` (bool, optional): If `True`, injects dynamic page numbers. Default: `False`.
*(Note: It has no `orientation` parameter).*

**Returns:**
- `BuildDocumentResult`: Structured generation summary.

---

### `batch_add_content`

```python
batch_add_content(elements: list[dict[str, typing.Any]], filename: str | None = None, file_path: str | None = None) -> dict[str, typing.Any]
```

Append multiple structured elements (headings, paragraphs, tables, lists) to an existing Word document in 1 call without reloading between steps.

**Parameters:**

- `filename` (str, optional): Target existing `.docx` file path (accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `elements` (list of dict): List of elements to append (same format as `build_document`).

**Returns:**
- `BuildDocumentResult`: Structured summary including `status`, `filename`, `total_elements`, `element_counts`, `summary`, and `message`.

---

## Document Tools

### `create_document`

```python
create_document(filename: str, title: str | None = None, author: str | None = None) -> dict[str, typing.Any]
```

Create a new Word document with optional metadata.

**Parameters:**

- `filename` (str): Name of the document to create (with or without .docx extension)
- `title` (str, optional): Document title
- `author` (str, optional): Document author

**Returns:**

- dict: Contains `status` ("success" or "error") and `message`

---

### `copy_document`

```python
copy_document(source_filename: str, destination_filename: str | None = None) -> dict[str, typing.Any]
```

Create a copy of an existing Word document.

**Parameters:**

- `source_filename` (str): Path to the source document
- `destination_filename` (str, optional): Path for the copy (default: generates a new name)

**Returns:**

- dict: Contains `status` and `message`

---

### `get_document_info`

```python
get_document_info(filename: str, response_format: str = 'markdown') -> dict[str, typing.Any]
```

Get metadata and properties of a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `response_format` (str, optional): Format of response - "markdown" for human-readable or "json" for structured data (default: "markdown")

**Returns:**

- dict: Returns `{"status", "properties", "response_format"}`. `properties` contains `author`, `title`, `subject`, `keywords`, `comments`, `category`, `last_modified_by`, `created`, `modified`, `filename`, `path`, `size_kb`. Page count is not computed; the OpenXML package does not expose a reliable page total without rendering.

---

### `get_document_text`

```python
get_document_text(filename: str, response_format: str = 'markdown') -> dict[str, typing.Any]
```

Extract all text from a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `response_format` (str, optional): Format of response - "markdown" for human-readable or "json" for structured data (default: "markdown")

**Returns:**

- dict: Contains `status` and `text` (returns `{"status", "text"}`).

---

### `get_document_outline`

```python
get_document_outline(filename: str, response_format: str = 'markdown') -> dict[str, typing.Any]
```

Get the structure/outline of a Word document including paragraphs and tables.

**Parameters:**

- `filename` (str): Path to the Word document
- `response_format` (str, optional): Format of response - "markdown" for human-readable or "json" for structured data (default: "markdown")

**Returns:**

- dict: Returns `paragraphs` and `tables` (`{"paragraphs": [{"index", "text", "style"}], "tables": [{"index", "rows", "columns", "preview"}]}`). Paragraph text is truncated to 100 characters plus `...`. Tables include a 3x3 cell preview. There is no heading tree and no `status` key on success.

---

### `list_available_documents`

```python
list_available_documents(directory: str | None = None, page: int = 1, page_size: int = 20, response_format: str = 'markdown') -> dict[str, typing.Any]
```

List all .docx files in the specified or allowed directories with pagination support. Only top-level `.docx` files in allowed directories are scanned.

**Parameters:**

- `directory` (str, optional): Directory to search (default: all allowed directories)
- `page` (int, optional): Page number for pagination (1-based, default: 1)
- `page_size` (int, optional): Number of documents per page (default: 20, max: 100)
- `response_format` (str, optional): Format of response - "markdown" or "json" (default: "markdown")

**Returns:**

- dict: Contains:
  - `status`: Operation status
  - `message`: Human-readable message
  - `documents`: List of documents with name, path, size_kb, source_directory
  - `pagination`: Object with `page`, `page_size`, `total`, `has_more`

---

### `merge_documents`

```python
merge_documents(target_filename: str, source_filenames: list[str], add_page_breaks: bool = True, response_format: str = 'markdown') -> dict[str, typing.Any]
```

Merge multiple Word documents into a single document.

**Parameters:**

- `target_filename` (str): Path to the target document
- `source_filenames` (list[str]): List of source document paths to merge
- `add_page_breaks` (bool, optional): Add page breaks between documents (default: True)
- `response_format` (str, optional): Format of response - "markdown" or "json" (default: "markdown")

**Returns:**

- dict: Contains `status`, `message`, and `file_path`.

---

## Content Tools

### `add_paragraph`

```python
add_paragraph(filename: str, text: str, style: str | None = None) -> dict[str, typing.Any]
```

Add a paragraph to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `text` (str): Text content for the paragraph
- `style` (str, optional): Paragraph style name

**Returns:**

- dict: Contains `status` and `message`

---

### `add_heading`

```python
add_heading(filename: str, text: str, level: int = 1) -> dict[str, typing.Any]
```

Add a heading to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `text` (str): Heading text
- `level` (int, optional): Heading level (1-9, where 1 is highest, default: 1)

**Returns:**

- dict: Contains `status`, `message`, and heading details

---

### `add_page_break`

```python
add_page_break(filename: str) -> dict[str, typing.Any]
```

Add a page break to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- dict: Contains `status` and `message`

---

### `add_picture`

```python
add_picture(filename: str, image_path: str, width: float | None = None) -> dict[str, typing.Any]
```

Add an image to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `image_path` (str): Path to the image file
- `width` (float, optional): Image width in inches (maintains aspect ratio)

**Returns:**

- dict: Contains `status`, `message`, and image details

---

### `add_table`

```python
add_table(filename: str, rows: int, cols: int, data: list[list[typing.Any]] | None = None) -> dict[str, typing.Any]
```

Add a table to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `rows` (int): Number of rows
- `cols` (int): Number of columns
- `data` (list[list[str]], optional): 2D list of cell values

**Returns:**

- dict: Contains `status`, `message`, table dimensions, and row count

---

### `add_table_of_contents`

```python
add_table_of_contents(filename: str, title: str = 'Table of Contents', max_level: int = 3) -> dict[str, typing.Any]
```

Add a table of contents to a Word document based on heading styles.

**Parameters:**

- `filename` (str): Path to the Word document
- `title` (str, optional): TOC title (default: "Table of Contents")
- `max_level` (int, optional): Maximum heading level to include (1-9, default: 3)

**Returns:**

- dict: Contains `status`, `message`, and entry count

---

### `delete_paragraph`

```python
delete_paragraph(filename: str, paragraph_index: int) -> dict[str, typing.Any]
```

Delete a paragraph from a Word document by index.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph to delete (0-based)

**Returns:**

- dict: Contains `status` and `message`

---

### `search_and_replace`

```python
search_and_replace(filename: str, find_text: str, replace_text: str) -> dict[str, typing.Any]
```

Search for text and replace all occurrences in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `find_text` (str): Text to search for
- `replace_text` (str): Text to replace with

**Returns:**

- dict: Contains `status`, `message`, and number of replacements made

---

## Format Tools

### `format_text`

```python
format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int, bold: bool | None = None, italic: bool | None = None, underline: bool | None = None, color: str | None = None, font_size: int | None = None, font_name: str | None = None) -> dict[str, typing.Any]
```

Format a specific range of text within a paragraph.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `start_pos` (int): Start position in the paragraph
- `end_pos` (int): End position in the paragraph
- `bold` (bool, optional): Apply bold formatting
- `italic` (bool, optional): Apply italic formatting
- `underline` (bool, optional): Apply underline
- `color` (str, optional): Text color (e.g., "red", "blue", "green")
- `font_size` (int, optional): Font size in points
- `font_name` (str, optional): Font name/family

**Returns:**

- dict: Contains `status`, `message`, and formatting details

---

### `create_custom_style`

```python
create_custom_style(filename: str, style_name: str, bold: bool | None = None, italic: bool | None = None, font_size: int | None = None, font_name: str | None = None, color: str | None = None, base_style: str | None = None) -> dict[str, typing.Any]
```

Create a custom style in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `style_name` (str): Name for the new style
- `bold` (bool, optional): Set text bold
- `italic` (bool, optional): Set text italic
- `font_size` (int, optional): Font size in points
- `font_name` (str, optional): Font name/family
- `color` (str, optional): Text color
- `base_style` (str, optional): Existing style to base this on

**Returns:**

- dict: Contains `status`, `message`, and style details

---

### `format_table`

```python
format_table(filename: str, table_index: int, has_header_row: bool | None = None, border_style: str | None = None, shading: list[list[str]] | None = None) -> dict[str, typing.Any]
```

Format a table with borders, shading, and structure options.

**Parameters:**

- `filename` (str): Path to the Word document
- `table_index` (int): Index of the table (0-based)
- `has_header_row` (bool, optional): Format first row as header
- `border_style` (str, optional): Border style ("none", "single", "double", "thick")
- `shading` (list[list[str]], optional): 2D list of cell background colors

**Returns:**

- dict: Contains `status`, `message`, and formatting details

---

## Protection Tools

### `protect_document`

```python
protect_document(filename: str, password: str) -> dict[str, typing.Any]
```

Add password protection to a Word document using encryption.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password to protect the document with

**Returns:**

- dict: Contains `status` and `message`

---

### `unprotect_document`

```python
unprotect_document(filename: str, password: str) -> dict[str, typing.Any]
```

Remove password protection from a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password that was used to protect the document

**Returns:**

- dict: Contains `status` and `message`

---

### `add_restricted_editing`

```python
add_restricted_editing(filename: str, password: str, editable_sections: list[str]) -> dict[str, typing.Any]
```

Add restricted editing to allow editing only in specified sections.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password to protect the document with
- `editable_sections` (list[str]): List of section names that can be edited (stored in the sidecar `.protection` file and is advisory)

**Returns:**

- dict: Contains `status`, `message`, and `editable_sections` list.

---

### `add_digital_signature`

```python
add_digital_signature(filename: str, signer_name: str, reason: str | None = None) -> dict[str, typing.Any]
```

Add a digital signature to a Word document. Writes a visible signature block and stores a SHA-256 content hash in the sidecar; it is not a cryptographic signature and does not verify signer identity.

**Parameters:**

- `filename` (str): Path to the Word document
- `signer_name` (str): Name of the person signing
- `reason` (str, optional): Reason for signing

**Returns:**

- dict: Contains `status`, `message`, and signature details

---

### `verify_document`

```python
verify_document(filename: str, password: str | None = None) -> dict[str, typing.Any]
```

Verify document protection and/or digital signature. It reads the sidecar file only.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str, optional): Password if document is protected

**Returns:**

- `ProtectionResult`: Structured result with `status`, `filename`, `protection_type`, `success`, and `message`.

---

## Footnote Tools

### `add_footnote_to_document`

```python
add_footnote_to_document(filename: str, paragraph_index: int, footnote_text: str) -> dict[str, typing.Any]
```

Add a footnote to a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `footnote_text` (str): Text content of the footnote

**Returns:**

- dict: Contains `status`, `message`, and footnote ID

---

### `add_endnote_to_document`

```python
add_endnote_to_document(filename: str, paragraph_index: int, endnote_text: str) -> dict[str, typing.Any]
```

Add an endnote to a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `endnote_text` (str): Text content of the endnote

**Returns:**

- dict: Contains `status`, `message`, and endnote ID

---

### `convert_footnotes_to_endnotes_in_document`

```python
convert_footnotes_to_endnotes_in_document(filename: str) -> dict[str, typing.Any]
```

Convert all footnotes to endnotes in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- dict: Contains `status`, `message`, and count of converted notes

---

### `customize_footnote_style`

```python
customize_footnote_style(filename: str, numbering_format: str = '1, 2, 3', start_number: int = 1, font_name: str | None = None, font_size: int | None = None) -> dict[str, typing.Any]
```

Customize the appearance of footnotes in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `numbering_format` (str, optional): Number format (default: `"1, 2, 3"`)
- `start_number` (int, optional): Starting number for footnotes (default: `1`)
- `font_size` (int, optional): Font size in points
- `font_name` (str, optional): Font name

**Returns:**

- dict: Contains `status`, `message`, and applied settings

---

## Extended Document Inspection

### `get_paragraph_text_from_document`

```python
get_paragraph_text_from_document(filename: str, paragraph_index: int) -> dict[str, typing.Any]
```

Get text from a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)

**Returns:**

- dict: Contains `status` and `paragraph` object with `index`, `text`, `style`, and `is_heading` (returns `{"status", "paragraph": {"index", "text", "style", "is_heading"}}`).

---

### `find_text_in_document`

```python
find_text_in_document(filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False) -> dict[str, typing.Any]
```

Search for text in a Word document and return matching locations.

**Parameters:**

- `filename` (str): Path to the Word document
- `text_to_find` (str): Text to search for
- `match_case` (bool, optional): Case-sensitive search (default: True)
- `whole_word` (bool, optional): Match whole words only (default: False)

**Returns:**

- dict: Returns `{"status", "search_results": {"query", "match_case", "whole_word", "occurrences", "total_count"}}` with `search_results`. Each occurrence has `location`, `position`, `match`, and `context`.

---

### `convert_to_pdf`

```python
convert_to_pdf(filename: str, output_filename: str | None = None) -> dict[str, typing.Any]
```

Convert a Word document to PDF format.

**Parameters:**

- `filename` (str): Path to the Word document
- `output_filename` (str, optional): Output path for the PDF (default: same as input with .pdf extension)

**Returns:**

- dict: Contains `status`, `pdf_path`, and `message`.

---

## Header & Footer

### `add_header`

```python
add_header(filename: str = '', text: str = '', section_index: int = 0, file_path: str | None = None) -> dict[str, typing.Any]
```

Add or update the primary header of a specific section in the document.

**Parameters:**

- `filename` (str): Output document or target file path (primary parameter, accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `text` (str): Text content for the header
- `section_index` (int, optional): Index of the section (0-based, default: 0)

**Returns:**

- dict: Contains `status`, `message`, and `file_path`

---

### `add_footer`

```python
add_footer(filename: str = '', text: str = '', section_index: int = 0, file_path: str | None = None) -> dict[str, typing.Any]
```

Add or update the primary footer of a specific section in the document.

**Parameters:**

- `filename` (str): Output document or target file path (primary parameter, accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `text` (str): Text content for the footer
- `section_index` (int, optional): Index of the section (0-based, default: 0)

**Returns:**

- dict: Contains `status`, `message`, and `file_path`

---

## Hyperlinks & Bookmarks

### `add_hyperlink`

```python
add_hyperlink(filename: str = '', paragraph_index: int = 0, text: str = '', url: str | None = None, bookmark: str | None = None, file_path: str | None = None) -> dict[str, typing.Any]
```

Add a hyperlink to a specific paragraph.

**Parameters:**

- `filename` (str): Target document file path (primary parameter, accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `paragraph_index` (int): Index of the paragraph (0-based)
- `text` (str): Display text for the hyperlink
- `url` (str, optional): Target URL (required if bookmark is not provided)
- `bookmark` (str, optional): Target bookmark name (required if url is not provided)

**Returns:**

- dict: Contains `status`, `message`, and `file_path`

---

### `add_bookmark`

```python
add_bookmark(filename: str = '', paragraph_index: int = 0, name: str = '', file_path: str | None = None) -> dict[str, typing.Any]
```

Add a bookmark to a specific paragraph to enable internal linking.

**Parameters:**

- `filename` (str): Target document file path (primary parameter, accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `paragraph_index` (int): Index of the paragraph (0-based)
- `name` (str): Name of the bookmark

**Returns:**

- dict: Contains `status`, `message`, and `file_path`

---

## Property & Layout Tools

### `get_core_properties`

```python
get_core_properties(file_path: str = '', filename: str | None = None) -> dict[str, typing.Any]
```

Get the core metadata properties of the document (author, title, etc.).

**Parameters:**

- `filename` (str, optional): Target document file path (accepts `filename` or `file_path`).
- `file_path` (str): Path to the Word document

**Returns:**

- dict: Contains `status` and `properties` object

---

### `set_core_properties`

```python
set_core_properties(file_path: str = '', author: str | None = None, title: str | None = None, subject: str | None = None, keywords: str | None = None, comments: str | None = None, category: str | None = None, filename: str | None = None) -> dict[str, typing.Any]
```

Set the core metadata properties of the document.

**Parameters:**

- `filename` (str, optional): Target document file path (accepts `filename` or `file_path`).
- `file_path` (str): Path to the Word document
- `author` (str, optional): Document author
- `title` (str, optional): Document title
- `subject` (str, optional): Document subject
- `keywords` (str, optional): Document keywords
- `comments` (str, optional): Document comments
- `category` (str, optional): Document category

**Returns:**

- dict: Contains `status`, `message`, and `file_path`

---

### `set_page_layout`

```python
set_page_layout(filename: str = '', section_index: int = 0, orientation: str = 'portrait', file_path: str | None = None) -> dict[str, typing.Any]
```

Set the page layout orientation (portrait or landscape) for a specific section.

**Parameters:**

- `filename` (str): Target document file path (primary parameter, accepts `filename` or `file_path`).
- `file_path` (str, optional): Alias for `filename`.
- `section_index` (int, optional): Index of the section (0-based, default: 0)
- `orientation` (str, optional): Page orientation ("portrait" or "landscape", default: "portrait")

**Returns:**

- dict: Contains `status`, `message`, and `file_path`

---

## Response Format

The `response_format` parameter (`"markdown"` or `"json"`) is declared by `get_document_info`, `get_document_text`, `get_document_outline`, `list_available_documents`, and `merge_documents`. Only `get_document_info` and `list_available_documents` echo the value in the response, and it does not change the output shape.

- `"markdown"` (default): Human-readable formatted text
- `"json"`: Machine-readable structured data

Example:

```json
{
  "status": "success",
  "message": "Document created successfully",
  "response_format": "json"
}
```

---

## Error Handling

All tools return a consistent error format:

```json
{
  "status": "error",
  "message": "Descriptive error message",
  "error_type": "FileNotFoundError",
  "suggestion": "Suggested solution if available",
  "recoverable": true
}
```

---

## Pagination

The `list_available_documents` tool supports pagination:

```json
{
  "status": "success",
  "documents": [...],
  "pagination": {
    "page": 1,
    "page_size": 20,
    "total": 45,
    "has_more": true
  }
}
```

To get the next page, call the tool again with `page = 2`.

---

*Last updated: November 2026*
*Version: 2.0.2*
