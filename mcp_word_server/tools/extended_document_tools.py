"""
Extended document tools for Word Document Server.

These tools provide enhanced document content extraction and search capabilities.
"""

# modulos estandar
import json
import os
import platform
import shutil
import subprocess
from typing import Any, Dict, Optional

# modulos propios
from mcp_word_server.utils.extended_document_utils import find_text, get_paragraph_text
from mcp_word_server.validation.document_validators import (
    check_file_writeable,
    validate_docx_file,
)


@validate_docx_file("filename")
async def get_paragraph_text_from_document(filename: str, paragraph_index: int) -> str:
    """Get text from a specific paragraph in a Word document.

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to retrieve (0-based)
    """
    if paragraph_index < 0:
        return "Invalid parameter: paragraph_index must be a non-negative integer"

    try:
        result = get_paragraph_text(filename, paragraph_index)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to get paragraph text: {str(e)}"


@validate_docx_file("filename")
async def find_text_in_document(
    filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False
) -> str:
    """Find occurrences of specific text in a Word document.

    Args:
        filename: Path to the Word document
        text_to_find: Text to search for in the document
        match_case: Whether to match case (True) or ignore case (False)
        whole_word: Whether to match whole words only (True) or substrings (False)
    """
    if not text_to_find:
        return "Search text cannot be empty"

    try:

        result = find_text(filename, text_to_find, match_case, whole_word)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to search for text: {str(e)}"


@check_file_writeable("filename")
@validate_docx_file("filename")
async def convert_to_pdf(
    filename: str, output_filename: Optional[str] = None
) -> Dict[str, Any]:
    """Convert a Word document to PDF format.

    Args:
        filename: Path to the Word document
        output_filename: Optional path for the output PDF. If not provided,
                         will use the same name with .pdf extension
    """
    # Generate output filename if not provided
    if not output_filename:
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}.pdf"
    elif not output_filename.lower().endswith(".pdf"):
        output_filename = f"{output_filename}.pdf"

    # Convert to absolute path if not already
    if not os.path.isabs(output_filename):
        output_filename = os.path.abspath(output_filename)

    # Ensure the output directory exists
    output_dir = os.path.dirname(output_filename)
    if not output_dir:
        output_dir = os.path.abspath(".")

    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    try:
        # Determine platform for appropriate conversion method
        system = platform.system()

        if system == "Windows":
            # On Windows, try docx2pdf which uses Microsoft Word
            try:
                from docx2pdf import convert

                convert(filename, output_filename)
                return {
                    "success": True,
                    "message": "Document successfully converted to PDF",
                    "pdf_path": output_filename,
                }
            except (ImportError, Exception) as e:
                return {
                    "success": False,
                    "error": str(e),
                    "hint": "docx2pdf requires Microsoft Word to be installed.",
                }
        elif system in ["Linux", "Darwin"]:  # Linux or macOS
            # Try using LibreOffice if available (common on Linux/macOS)
            try:
                # Choose the appropriate command based on OS
                if system == "Darwin":  # macOS
                    lo_commands = [
                        "soffice",
                        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                    ]
                else:  # Linux
                    lo_commands = ["libreoffice", "soffice"]

                # Try each possible command
                conversion_successful = False
                errors = []

                for cmd_name in lo_commands:
                    try:
                        # Construct LibreOffice conversion command
                        output_dir = os.path.dirname(output_filename)
                        # If output_dir is empty, use current directory
                        if not output_dir:
                            output_dir = "."
                        # Ensure the directory exists
                        os.makedirs(output_dir, exist_ok=True)

                        cmd = [
                            cmd_name,
                            "--headless",
                            "--convert-to",
                            "pdf",
                            "--outdir",
                            output_dir,
                            filename,
                        ]

                        result = subprocess.run(
                            cmd, capture_output=True, text=True, timeout=60
                        )

                        if result.returncode == 0:
                            # LibreOffice creates the PDF with the same basename
                            base_name = os.path.basename(filename)
                            pdf_base_name = os.path.splitext(base_name)[0] + ".pdf"
                            created_pdf = os.path.join(
                                os.path.dirname(output_filename) or ".", pdf_base_name
                            )

                            # If the created PDF is not at the desired location, move it
                            if created_pdf != output_filename and os.path.exists(
                                created_pdf
                            ):
                                shutil.move(created_pdf, output_filename)

                            conversion_successful = True
                            break  # Exit the loop if successful
                        else:
                            errors.append(f"{cmd_name} error: {result.stderr}")
                    except (subprocess.SubprocessError, FileNotFoundError) as e:
                        errors.append(f"{cmd_name} error: {str(e)}")

                if conversion_successful:
                    return {
                        "success": True,
                        "message": "Document successfully converted to PDF",
                        "pdf_path": output_filename,
                    }
                else:
                    # If all LibreOffice attempts failed, try docx2pdf as fallback
                    try:
                        from docx2pdf import convert

                        convert(filename, output_filename)
                        return {
                            "success": True,
                            "message": "Document converted using fallback docx2pdf",
                            "pdf_path": output_filename,
                        }
                    except (ImportError, Exception) as e:
                        return {
                            "success": False,
                            "error": "Conversion failed using both LibreOffice and docx2pdf",
                            "details": {
                                "libreoffice_errors": errors,
                                "docx2pdf_error": str(e),
                            },
                            "hint": "Install LibreOffice or Microsoft Word depending on your OS.",
                        }

            except Exception as e:
                return {
                    "success": False,
                    "error": f"Failed to convert document to PDF: {str(e)}",
                }
        else:
            return {"success": False, "error": f"Unsupported platform: {system}"}

    except Exception as e:
        return {"success": False, "error": str(e)}
