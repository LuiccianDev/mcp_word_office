"""Document context management for Word Document Server.

This module provides a context manager to handle opening and saving Word documents,
ensuring consistent error handling and reducing redundant file operations.
"""

import os
from contextlib import contextmanager
from typing import Generator

from docx import Document
from docx.document import Document as DocumentType
from docx.opc.exceptions import PackageNotFoundError

from mcp_word.exception import DocumentProcessingError


@contextmanager
def document_context(
    filename: str, mode: str = "read"
) -> Generator[DocumentType, None, None]:
    """Context manager for handling Word document lifecycle.

    Args:
        filename: Path to the .docx file.
        mode: Operation mode ('read' or 'write'). Default is 'read'.

    Yields:
        DocumentType: The loaded docx.Document object.

    Raises:
        DocumentProcessingError: For any document-related failures.
    """
    if not os.path.exists(filename) and mode == "read":
        raise DocumentProcessingError(f"Document not found: {filename}")

    try:
        if os.path.exists(filename):
            doc = Document(filename)
        else:
            # For new documents, though usually handled by create_document tool
            doc = Document()

        yield doc

        if mode == "write":
            doc.save(filename)

    except PackageNotFoundError:
        raise DocumentProcessingError(f"Invalid Word document: {filename}")
    except OSError as error:
        raise DocumentProcessingError(f"File system error: {str(error)}")
    except Exception as error:
        if not isinstance(error, DocumentProcessingError):
            raise DocumentProcessingError(f"Unexpected error: {str(error)}") from error
        raise
