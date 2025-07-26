# MCP Office Word - Tools Reference

This document provides a comprehensive reference for all tools available in the MCP Office Word server. These tools allow you to create, modify, and manage Word documents programmatically.

## Table of Contents

- [Document Tools](#document-tools)
- [Content Tools](#content-tools)
- [Format Tools](#format-tools)
- [Protection Tools](#protection-tools)
- [Footnote Tools](#footnote-tools)
- [Extended Document Tools](#extended-document-tools)

## Document Tools

### `create_document`

Create a new Word document with optional metadata.

**Parameters:**

- `filename` (str): Name of the document to create (with or without .docx extension)
- `title` (str, optional): Document title
- `author` (str, optional): Document author

**Returns:**

- str: Success or error message

---

### `copy_document`

Create a copy of an existing Word document.

**Parameters:**

- `source_filename` (str): Path to the source document
- `destination_filename` (str, optional): Path for the copy (default: generates a new name)

**Returns:**

- str: Success or error message

---

### `get_document_info`

Get metadata and properties of a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- str: JSON string containing document properties

---

### `get_document_text`

Extract all text from a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- str: Extracted text content

---

### `get_document_outline`

Get the structure/outline of a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- str: JSON string containing document structure

---

### `list_available_documents`

List all .docx files in the specified or allowed directories.

**Parameters:**

- `directory` (str, optional): Directory to search (default: all allowed directories)

**Returns:**

- dict: Dictionary containing list of documents and search information

---

### `merge_documents`

Merge multiple Word documents into a single document.

**Parameters:**

- `target_filename` (str): Path to the target document
- `source_filenames` (List[str]): List of source document paths to merge
- `add_page_breaks` (bool, optional): Add page breaks between documents (default: True)

**Returns:**

- str: Success or error message

## Content Tools

### `add_paragraph`

Add a paragraph to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `text` (str): Text content for the paragraph
- `style` (str, optional): Paragraph style name

**Returns:**

- str: Success or error message

---

### `add_heading`

Add a heading to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `text` (str): Heading text
- `level` (int, optional): Heading level (1-9, default: 1)

**Returns:**

- str: Success or error message

---

### `add_page_break`

Add a page break to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- str: Success or error message

---

### `add_picture`

Add an image to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `image_path` (str): Path to the image file
- `width` (float, optional): Image width in inches (maintains aspect ratio)

**Returns:**

- str: Success or error message

---

### `add_table`

Add a table to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `rows` (int): Number of rows
- `cols` (int): Number of columns
- `data` (List[List[str]], optional): 2D list of cell values
- `style` (str, optional): Table style name

**Returns:**

- str: Success or error message

---

### `add_table_of_contents`

Add a table of contents to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `title` (str, optional): TOC title (default: "Table of Contents")
- `max_level` (int, optional): Maximum heading level to include (1-9, default: 3)

**Returns:**

- str: Success or error message

---

### `search_and_replace`

Search for text and replace all occurrences in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `find_text` (str): Text to search for
- `replace_text` (str): Text to replace with

**Returns:**

- str: Success or error message

## Format Tools

### `format_text`

Format a specific range of text within a paragraph.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `start_pos` (int): Start position in the paragraph
- `end_pos` (int): End position in the paragraph
- `bold` (bool, optional): Apply bold formatting
- `italic` (bool, optional): Apply italic formatting
- `underline` (bool, optional): Apply underline
- `color` (str, optional): Text color (hex or color name)
- `font_size` (int, optional): Font size in points
- `font_name` (str, optional): Font name

**Returns:**

- str: Success or error message

---

### `format_table`

Format a table in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `table_index` (int): Index of the table (0-based)
- `style` (str, optional): Table style name
- `alignment` (str, optional): Table alignment ('left', 'center', 'right')
- `width` (float, optional): Table width in inches

**Returns:**

- str: Success or error message

## Protection Tools

### `protect_document`

Protect a Word document with a password.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str): Password for protection
- `protection_type` (str, optional): Type of protection ('read_only', 'comments', 'tracked_changes', 'forms')

**Returns:**

- str: Success or error message

---

### `unprotect_document`

Remove protection from a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `password` (str, optional): Password if document is protected

**Returns:**

- str: Success or error message

---

### `add_digital_signature`

Add a digital signature to a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `certificate_path` (str): Path to the certificate file
- `reason` (str, optional): Reason for signing
- `location` (str, optional): Location of signing

**Returns:**

- str: Success or error message

## Footnote Tools

### `add_footnote_to_document`

Add a footnote to a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `footnote_text` (str): Text content of the footnote

**Returns:**

- str: Success or error message

---

### `add_endnote_to_document`

Add an endnote to a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `endnote_text` (str): Text content of the endnote

**Returns:**

- str: Success or error message

---

### `convert_footnotes_to_endnotes_in_document`

Convert all footnotes to endnotes in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document

**Returns:**

- str: Success or error message

---

### `customize_footnote_style`

Customize the appearance of footnotes in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `number_format` (str, optional): Number format ('decimal', 'lowerRoman', 'upperRoman', 'lowerLetter', 'upperLetter')
- `start_at` (int, optional): Starting number for footnotes
- `restart_each_page` (bool, optional): Restart numbering on each page

**Returns:**

- str: Success or error message

## Extended Document Tools

### `get_paragraph_text_from_document`

Get text from a specific paragraph in a Word document.

**Parameters:**

- `filename` (str): Path to the Word document
- `paragraph_index` (int): Index of the paragraph (0-based)

**Returns:**

- str: Paragraph text or error message

---

### `find_text_in_document`

Search for text in a Word document and return matching locations.

**Parameters:**

- `filename` (str): Path to the Word document
- `search_text` (str): Text to search for
- `match_case` (bool, optional): Case-sensitive search (default: False)
- `match_whole_word` (bool, optional): Match whole words only (default: False)

**Returns:**

- str: JSON string containing match information or error message

---

### `convert_to_pdf`

Convert a Word document to PDF format.

**Parameters:**

- `filename` (str): Path to the Word document
- `output_path` (str, optional): Output path for the PDF (default: same as input with .pdf extension)

**Returns:**

- str: Success or error message
