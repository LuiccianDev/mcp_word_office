# MCP Word Office - Tools Reference

This document provides a comprehensive reference for all tools available in the MCP Word Office server. These tools allow you to create, modify, and manage Word documents programmatically.

> **Note:** All tools support two naming conventions:
> - **New naming** (recommended): Tools use the `word_` prefix (e.g., `word_create_document`)
> - **Legacy naming**: Original names without prefix for backward compatibility (e.g., `create_document`)

## Table of Contents

- [Document Tools](#document-tools)
- [Content Tools](#content-tools)
- [Format Tools](#format-tools)
- [Protection Tools](#protection-tools)
- [Footnote Tools](#footnote-tools)
- [Extended Document Tools](#extended-document-tools)

---

## Document Tools

### `word_create_document` / `create_document`

Create a new Word document with optional metadata.

**Parameters:**

- `filename` (str): Name of the document to create (with or without .docx extension)
- `title` (str, optional): Document title
- `author` (str, optional): Document author

**Returns:**

- dict: Contains `status` ("success" or "error") and `message`

---

### `word_copy_document` / `copy_document`

Create a copy of an existing Word document.

**Parameters:**

- `source_filename` (str): Path to the source document
- `destination_filename` (str, optional): Path for the copy (default: generates a new name)

**Returns:**

- dict: Contains `status` and `message`

---

### `word_get_document_info` / `get_document_info`

Get metadata and properties of a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `response_format` (str, optional): Format of response - "markdown" for human-readable or "json" for structured data (default: "markdown")

**Returns:**

- dict: Document properties including title, author, word_count, page_count, created/modified dates, etc.

---

### `word_get_document_text` / `get_document_text`

Extract all text from a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `response_format` (str, optional): Format of response - "markdown" for human-readable or "json" for structured data (default: "markdown")

**Returns:**

- dict: Contains `full_text`, `paragraph_count`, and `paragraphs` list

---

### `word_get_document_outline` / `get_document_outline`

Get the structure/outline of a Word document including headings and tables.

**Parameters:**

- `filename` (str): Path to the Word document
- `response_format` (str, optional): Format of response - "markdown" for human-readable or "json" for structured data (default: "markdown")

**Returns:**

- dict: Contains `headings` (list with level, text, position) and `tables` (list with index, row_count, column_count)

---

### `word_list_documents` / `list_available_documents`

List all .docx files in the specified or allowed directories with pagination support.

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
  - `pagination`: Object with page, page_size, total, has_more, next_offset
  - `directories_searched`: List of directories that were searched

---

### `word_merge_documents` / `merge_documents`

Merge multiple Word documents into a single document.

**Parameters:**

- `target_filename` (str): Path to the target document
- `source_filenames` (list[str]): List of source document paths to merge
- `add_page_breaks` (bool, optional): Add page breaks between documents (default: True)
- `response_format` (str, optional): Format of response - "markdown" or "json" (default: "markdown")

**Returns:**

- dict: Contains `status`, `message`, and details about merged documents

---

## Content Tools

### `word_add_paragraph` / `add_paragraph`

Add a paragraph to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `text` (str): Text content for the paragraph
- `style` (str, optional): Paragraph style name

**Returns:**

- dict: Contains `status` and `message`

---

### `word_add_heading` / `add_heading`

Add a heading to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `text` (str): Heading text
- `level` (int, optional): Heading level (1-9, where 1 is highest, default: 1)

**Returns:**

- dict: Contains `status`, `message`, and heading details

---

### `word_add_page_break` / `add_page_break`

Add a page break to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- dict: Contains `status` and `message`

---

### `word_add_picture` / `add_picture`

Add an image to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `image_path` (str): Path to the image file
- `width` (float, optional): Image width in inches (maintains aspect ratio)

**Returns:**

- dict: Contains `status`, `message`, and image details

---

### `word_add_table` / `add_table`

Add a table to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `rows` (int): Number of rows
- `cols` (int): Number of columns
- `data` (list[list[str]], optional): 2D list of cell values

**Returns:**

- dict: Contains `status`, `message`, table dimensions, and row count

---

### `word_add_table_of_contents` / `add_table_of_contents`

Add a table of contents to a Word document based on heading styles.

**Parameters:**

- `filename` (str): Path to the Word document
- `title` (str, optional): TOC title (default: "Table of Contents")
- `max_level` (int, optional): Maximum heading level to include (1-9, default: 3)

**Returns:**

- dict: Contains `status`, `message`, and entry count

---

### `word_delete_paragraph` / `delete_paragraph`

Delete a paragraph from a Word document by index.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph to delete (0-based)

**Returns:**

- dict: Contains `status` and `message`

---

### `word_search_and_replace` / `search_and_replace`

Search for text and replace all occurrences in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `find_text` (str): Text to search for
- `replace_text` (str): Text to replace with

**Returns:**

- dict: Contains `status`, `message`, and number of replacements made

---

## Format Tools

### `word_format_text` / `format_text`

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

### `word_create_custom_style` / `create_custom_style`

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

### `word_format_table` / `format_table`

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

### `word_protect_document` / `protect_document`

Add password protection to a Word document using encryption.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password to protect the document with

**Returns:**

- dict: Contains `status` and `message`

---

### `word_unprotect_document` / `unprotect_document`

Remove password protection from a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password that was used to protect the document

**Returns:**

- dict: Contains `status` and `message`

---

### `word_add_restricted_editing` / `add_restricted_editing`

Add restricted editing to allow editing only in specified sections.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password to protect the document with
- `editable_sections` (list[str]): List of section names that can be edited

**Returns:**

- dict: Contains `status`, `message`, and editable sections

---

### `word_add_digital_signature` / `add_digital_signature`

Add a digital signature to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `signer_name` (str): Name of the person signing
- `reason` (str, optional): Reason for signing

**Returns:**

- dict: Contains `status`, `message`, and signature details

---

### `word_verify_document` / `verify_document`

Verify document protection and/or digital signature.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str, optional): Password if document is protected

**Returns:**

- dict: Contains `status`, `message`, and verification details

---

## Footnote Tools

### `word_add_footnote` / `add_footnote_to_document`

Add a footnote to a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `footnote_text` (str): Text content of the footnote
- `reference_mark` (str, optional): Custom reference mark

**Returns:**

- dict: Contains `status`, `message`, and footnote ID

---

### `word_add_endnote` / `add_endnote_to_document`

Add an endnote to a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `endnote_text` (str): Text content of the endnote

**Returns:**

- dict: Contains `status`, `message`, and endnote ID

---

### `word_convert_footnotes` / `convert_footnotes_to_endnotes_in_document`

Convert all footnotes to endnotes in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- dict: Contains `status`, `message`, and count of converted notes

---

### `word_customize_footnote_style` / `customize_footnote_style`

Customize the appearance of footnotes in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `number_format` (str, optional): Number format ("decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter")
- `start_at` (int, optional): Starting number for footnotes
- `restart_each_page` (bool, optional): Restart numbering on each page
- `font_size` (int, optional): Font size in points
- `font_name` (str, optional): Font name

**Returns:**

- dict: Contains `status`, `message`, and applied settings

---

## Extended Document Tools

### `word_get_paragraph_text` / `get_paragraph_text_from_document`

Get text from a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)

**Returns:**

- dict: Contains `status`, `message`, `text`, and `character_count`

---

### `word_find_text` / `find_text_in_document`

Search for text in a Word document and return matching locations.

**Parameters:**

- `filename` (str): Path to the Word document
- `search_text` (str): Text to search for
- `match_case` (bool, optional): Case-sensitive search (default: True)
- `whole_word` (bool, optional): Match whole words only (default: False)

**Returns:**

- dict: Contains `status`, `message`, `match_count`, and `matches` list

---

### `word_convert_to_pdf` / `convert_to_pdf`

Convert a Word document to PDF format.

**Parameters:**

- `filename` (str): Path to the Word document
- `output_filename` (str, optional): Output path for the PDF (default: same as input with .pdf extension)

**Returns:**

- dict: Contains `status`, `success`, `pdf_path`, and `message`

---

## Response Format

All tools that return data support a `response_format` parameter:

- `"markdown"` (default): Human-readable formatted text with headers and lists
- `"json"`: Machine-readable structured data for programmatic processing

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

The `word_list_documents` tool supports pagination:

```json
{
  "status": "success",
  "documents": [...],
  "pagination": {
    "page": 1,
    "page_size": 20,
    "total": 45,
    "has_more": true,
    "next_offset": 2
  }
}
```

To get the next page, call the tool again with `page = 2`.

---

*Last updated: February 2026*
*Version: 1.0.1*
