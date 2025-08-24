"""
Formatting tools for Word Document Server.

These tools handle formatting operations for Word documents,
including text formatting, table formatting, and custom styles.
"""

# modulos estandar
from typing import Any

# modulos de terceros
from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor

from mcp_word.core.styles import create_style
from mcp_word.core.tables import apply_table_style

# modulos propios
from mcp_word.validation.document_validators import (
    check_file_writeable,
    validate_docx_file,
)


@validate_docx_file("filename")
@check_file_writeable("filename")
async def format_text(
    filename: str,
    paragraph_index: int,
    start_pos: int,
    end_pos: int,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    font_size: int | None = None,
    font_name: str | None = None,
) -> dict[str, Any]:
    """Format a specific range of text within a paragraph.

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph (0-based)
        start_pos: Start position within the paragraph text
        end_pos: End position within the paragraph text
        bold: Set text bold (True/False)
        italic: Set text italic (True/False)
        underline: Set text underlined (True/False)
        color: Text color (e.g., 'red', 'blue', etc.)
        font_size: Font size in points
        font_name: Font name/family
    """

    # Ensure numeric parameters are the correct type
    try:
        paragraph_index = int(paragraph_index)
        start_pos = int(start_pos)
        end_pos = int(end_pos)
        if font_size is not None:
            font_size = int(font_size)
    except (ValueError, TypeError):
        return {
            "status": "error",
            "error": "Invalid parameter: paragraph_index, start_pos, end_pos, and font_size must be integers",
        }

    try:
        doc: DocumentType = Document(filename)

        # Validate paragraph index
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return {
                "status": "error",
                "error": f"Invalid paragraph index. Document has {len(doc.paragraphs)} paragraphs (0-{len(doc.paragraphs) - 1}).",
            }

        paragraph = doc.paragraphs[paragraph_index]
        text = paragraph.text

        # Validate text positions
        if start_pos < 0 or end_pos > len(text) or start_pos >= end_pos:
            return {
                "status": "error",
                "error": f"Invalid text positions. Paragraph has {len(text)} characters.",
            }

        # Get the text to format
        target_text = text[start_pos:end_pos]

        # Clear existing runs and create three runs: before, target, after
        for run in paragraph.runs:
            run.clear()

        # Add text before target
        if start_pos > 0:
            paragraph.add_run(text[:start_pos])

        # Add target text with formatting
        run_target = paragraph.add_run(target_text)
        if bold is not None:
            run_target.bold = bold
        if italic is not None:
            run_target.italic = italic
        if underline is not None:
            run_target.underline = underline
        if color:
            # Define common RGB colors
            color_map = {
                "red": RGBColor(255, 0, 0),
                "blue": RGBColor(0, 0, 255),
                "green": RGBColor(0, 128, 0),
                "yellow": RGBColor(255, 255, 0),
                "black": RGBColor(0, 0, 0),
                "gray": RGBColor(128, 128, 128),
                "white": RGBColor(255, 255, 255),
                "purple": RGBColor(128, 0, 128),
                "orange": RGBColor(255, 165, 0),
            }

            try:
                if color.lower() in color_map:
                    # Use predefined RGB color
                    run_target.font.color.rgb = color_map[color.lower()]
                else:
                    # Try to set color by name
                    run_target.font.color.rgb = RGBColor.from_string(color)
            except Exception:
                # If all else fails, default to black
                run_target.font.color.rgb = RGBColor(0, 0, 0)
        if font_size:
            run_target.font.size = Pt(font_size)
        if font_name:
            run_target.font.name = font_name

        # Add text after target
        if end_pos < len(text):
            paragraph.add_run(text[end_pos:])

        doc.save(filename)
        return {
            "status": "success",
            "message": f"Text '{target_text}' formatted successfully in paragraph {paragraph_index}.",
        }
    except Exception as e:
        return {
            "status": "error",
            "error": f"Failed to format text: {str(e)}",
        }


@validate_docx_file("filename")
@check_file_writeable("filename")
async def create_custom_style(
    filename: str,
    style_name: str,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: int | None = None,
    font_name: str | None = None,
    color: str | None = None,
    base_style: str | None = None,
) -> dict[str, Any]:
    """Create a custom style in the document.

    Args:
        filename: Path to the Word document
        style_name: Name for the new style
        bold: Set text bold (True/False)
        italic: Set text italic (True/False)
        font_size: Font size in points
        font_name: Font name/family
        color: Text color (e.g., 'red', 'blue')
        base_style: Optional existing style to base this on
    """
    try:
        doc: DocumentType = Document(filename)

        # Build font properties dictionary
        font_properties: dict[str, Any] = {}
        if bold is not None:
            font_properties["bold"] = bold
        if italic is not None:
            font_properties["italic"] = italic
        if font_size is not None:
            font_properties["size"] = font_size
        if font_name is not None:
            font_properties["name"] = font_name
        if color is not None:
            font_properties["color"] = color

        # Create the style
        create_style(
            doc,
            style_name,
            WD_STYLE_TYPE.PARAGRAPH,
            base_style=base_style,
            font_props=font_properties,
        )

        doc.save(filename)
        return {
            "status": "success",
            "message": f"Style '{style_name}' created successfully.",
        }
    except Exception as e:
        return {
            "status": "error",
            "error": f"Failed to create style: {str(e)}",
        }


@validate_docx_file("filename")
@check_file_writeable("filename")
async def format_table(
    filename: str,
    table_index: int,
    has_header_row: bool | None = None,
    border_style: str | None = None,
    shading: list[list[str]] | None = None,
) -> dict[str, Any]:
    """Format a table with borders, shading, and structure.

    Args:
        filename: Path to the Word document
        table_index: Index of the table (0-based)
        has_header_row: If True, formats the first row as a header
        border_style: Style for borders ('none', 'single', 'double', 'thick')
        shading: 2D list of cell background colors (by row and column)
    """
    try:
        doc: DocumentType = Document(filename)

        # Validate table index
        if table_index < 0 or table_index >= len(doc.tables):
            return {
                "status": "error",
                "error": f"Invalid table index. Document has {len(doc.tables)} tables (0-{len(doc.tables) - 1}).",
            }

        table = doc.tables[table_index]

        # Apply formatting
        # Convert shading to the expected type
        table_shading: list[list[str | None]] | None = None
        if shading is not None:
            table_shading = [list(row) for row in shading]

        success = apply_table_style(
            table, has_header_row or False, border_style, table_shading
        )

        if success:
            doc.save(filename)
            return {
                "status": "success",
                "message": f"Table at index {table_index} formatted successfully.",
            }
        else:
            return {
                "status": "error",
                "error": f"Failed to format table at index {table_index}.",
            }
    except Exception as e:
        return {
            "status": "error",
            "error": f"Failed to format table: {str(e)}",
        }
