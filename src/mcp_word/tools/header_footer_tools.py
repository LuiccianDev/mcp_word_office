"""FastMCP tools for document headers and footers manipulation."""

from typing import Any

from docx import Document

from mcp_word.core import headers_footers
from mcp_word.validation.document_validators import validate_docx_write


@validate_docx_write("file_path")
async def add_header(file_path: str, text: str, section_index: int = 0) -> dict[str, Any]:
    """Add or update the primary header of a specific section in the document.

    Args:
        file_path: Path to the .docx file
        text: The text content for the header
        section_index: The index of the section (0-based). Default is 0.

    Returns:
        A dictionary containing the status and a message or error details
    """
    try:
        doc = Document(file_path)
        
        headers_footers.set_section_header(doc, text, section_idx=section_index)
        
        doc.save(file_path)
        
        return {
            "status": "success",
            "message": f"Header successfully updated in section {section_index} of {file_path}",
            "file_path": str(file_path)
        }
    except IndexError as e:
        return {"status": "error", "message": str(e)}
    except Exception as e:
        return {"status": "error", "message": f"Failed to add header: {e!s}"}


@validate_docx_write("file_path")
async def add_footer(file_path: str, text: str, section_index: int = 0) -> dict[str, Any]:
    """Add or update the primary footer of a specific section in the document.

    Args:
        file_path: Path to the .docx file
        text: The text content for the footer
        section_index: The index of the section (0-based). Default is 0.

    Returns:
        A dictionary containing the status and a message or error details
    """
    try:
        doc = Document(file_path)

        headers_footers.set_section_footer(doc, text, section_idx=section_index)

        doc.save(file_path)

        return {
            "status": "success",
            "message": f"Footer successfully updated in section {section_index} of {file_path}",
            "file_path": str(file_path)
        }
    except IndexError as e:
        return {"status": "error", "message": str(e)}
    except Exception as e:
        return {"status": "error", "message": f"Failed to add footer: {e!s}"}
