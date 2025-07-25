"""Document creation and manipulation tools for Word Document Server.

This module provides functions for creating, reading, and manipulating
Word documents with security checks for file system access.
"""
# modulos estandar
import json
import os
from typing import List, Optional, Tuple

# modulos de terceros
from docx import Document
from docx.document import Document as DocumentType

# modulos propios
from mcp_word_server.core.tables import copy_table
from mcp_word_server.utils.document_utils import (get_document_properties, 
                                                    extract_document_text, 
                                                    get_document_structure,
                                                    create_document_copy,
                                                    ensure_docx_extension)
from mcp_word_server.core.styles import (ensure_heading_style,
                                        ensure_table_style)
from mcp_word_server.validation.document_validators import (validate_docx_file,
                                                            check_file_writeable)


def _get_allowed_directories() -> List[str]:
    """Get the list of allowed directories from environment variables.
    
    Returns:
        List of absolute paths to directories where documents can be created/accessed.
        Defaults to ['./documents'] if MCP_ALLOWED_DIRECTORIES is not set.
    """
    
    allowed_dirs_str = os.environ.get("MCP_ALLOWED_DIRECTORIES", "./documents")
    
    allowed_dirs = [dir.strip() for dir in allowed_dirs_str.split(",")]
    
    return [os.path.abspath(dir) for dir in allowed_dirs]

def _is_path_in_allowed_directories(file_path: str) -> Tuple[bool, Optional[str]]:
    """Check if the given file path is within allowed directories.
    
    Args:
        file_path: The file path to validate.
        
    Returns:
        Tuple of (is_allowed, error_message) where is_allowed is a boolean
        indicating if the path is allowed, and error_message provides details
        if the path is not allowed.
    """
    allowed_dirs = _get_allowed_directories()
    abs_path = os.path.abspath(file_path)
    
    for allowed_dir in allowed_dirs:
        try:
            if os.path.commonpath([allowed_dir, abs_path]) == allowed_dir:
                return True, None
        except ValueError:
            continue
    
    error_msg = (
        f"Path '{file_path}' is not in allowed directories: "
        f"{', '.join(allowed_dirs)}"
    )
    return False, error_msg


async def create_document(
    filename: str, 
    title: Optional[str] = None, 
    author: Optional[str] = None
) -> str:
    """Create a new Word document with optional metadata.
    
    Args:
        filename: Name of the document to create (with or without .docx extension).
        title: Optional title for the document metadata.
        author: Optional author for the document metadata.
        
    Returns:
        str: Success or error message indicating the result of the operation.
    """

    is_allowed, error_message = _is_path_in_allowed_directories(filename)
    if not is_allowed:
        return f"Cannot create document: {error_message}"

    directory = os.path.dirname(filename)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as e:
            return f"Cannot create directory '{directory}': {str(e)}"
    
    try:
        doc: DocumentType = Document()
        
        # Set properties if provided
        if title:
            doc.core_properties.title = title
        if author:
            doc.core_properties.author = author
        
        # Ensure necessary styles exist
        ensure_heading_style(doc)
        ensure_table_style(doc)
        
        # Save the document
        doc.save(filename)
        
        return f"Document {filename} created successfully"
    except Exception as e:
        return f"Failed to create document: {str(e)}"


#! Function removed 
#async def create_document(filename: str, title: Optional[str] = None, author: Optional[str] = None) -> str:
    """Create a new Word document with optional metadata.
    
    Args:
        filename: Name of the document to create (with or without .docx extension)
        title: Optional title for the document metadata
        author: Optional author for the document metadata
    """
    """ filename = ensure_docx_extension(filename)
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot create document: {error_message}"
    
    try:
        doc = Document()
        
        # Set properties if provided
        if title:
            doc.core_properties.title = title
        if author:
            doc.core_properties.author = author
        
        # Ensure necessary styles exist
        ensure_heading_style(doc)
        ensure_table_style(doc)
        
        # Save the document
        doc.save(filename)
        
        return f"Document {filename} created successfully"
    except Exception as e:
        return f"Failed to create document: {str(e)}" """


@validate_docx_file('filename')
async def get_document_info(filename: str) -> str:
    """Get information about a Word document.
    
    Args:
        filename: Path to the Word document
    """
    try:
        properties = get_document_properties(filename)
        return json.dumps(properties, indent=2)
    except Exception as e:
        return f"Failed to get document info: {str(e)}"


@validate_docx_file('filename')
async def get_document_text(filename: str) -> str:
    """Extract all text from a Word document.
    
    Args:
        filename: Path to the Word document
    """ 
    return extract_document_text(filename)


@validate_docx_file('filename')
async def get_document_outline(filename: str) -> str:
    """Get the structure of a Word document.
    
    Args:
        filename: Path to the Word document
    """
    structure = get_document_structure(filename)
    return json.dumps(structure, indent=2)


async def list_available_documents(directory: str = ".") -> str:
    """List all .docx files in the specified directory.
    
    Args:
        directory: Directory to search for Word documents
    """
    try:
        if not os.path.exists(directory):
            return f"Directory {directory} does not exist"
        
        docx_files = [
            f for f in os.listdir(directory) 
            if f.lower().endswith('.docx')
        ]
        
        if not docx_files:
            return f"No Word documents found in {directory}"
        
        result = [f"Found {len(docx_files)} Word documents in {directory}:"]
        for file in sorted(docx_files):
            file_path = os.path.join(directory, file)
            size_kb = os.path.getsize(file_path) / 1024
            result.append(f"- {file} ({size_kb:.2f} KB)")
        
        return "\n".join(result)
    except Exception as e:
        return f"Failed to list documents: {str(e)}"

@validate_docx_file('source_filename')
async def copy_document(source_filename: str, destination_filename: Optional[str] = None) -> str:
    """Create a copy of a Word document.
    
    Args:
        source_filename: Path to the source document
        destination_filename: Optional path for the copy. If not provided, a default name will be generated.
    """
    source_filename = ensure_docx_extension(source_filename)
    
    if destination_filename:
        destination_filename = ensure_docx_extension(destination_filename)
    
    success, message, _ = create_document_copy(source_filename, destination_filename)
    if success:
        return message
    return f"Failed to copy document: {message}"

@check_file_writeable('target_filename')
@validate_docx_file('target_filename')
async def merge_documents(target_filename: str, source_filenames: List[str], add_page_breaks: bool = True) -> str:
    """Merge multiple Word documents into a single document.
    
    Args:
        target_filename: Path to the target document (will be created or overwritten)
        source_filenames: List of paths to source documents to merge
        add_page_breaks: If True, add page breaks between documents
    """
    # Validate all source documents exist
    missing_files = []
    for filename in source_filenames:
        doc_filename = ensure_docx_extension(filename)
        if not os.path.exists(doc_filename):
            missing_files.append(doc_filename)
    
    if missing_files:
        return f"Cannot merge documents. The following source files do not exist: {', '.join(missing_files)}"
    
    try:
        # Create a new document for the merged result
        target_doc: DocumentType = Document()
        
        # Process each source document
        for i, filename in enumerate(source_filenames):
            doc_filename = ensure_docx_extension(filename)
            source_doc: DocumentType = Document(doc_filename)
            
            # Add page break between documents (except before the first one)
            if add_page_breaks and i > 0:
                target_doc.add_page_break()
            
            # Copy all paragraphs
            for paragraph in source_doc.paragraphs:
                # Create a new paragraph with the same text and style
                new_paragraph = target_doc.add_paragraph(paragraph.text)
                new_paragraph.style = target_doc.styles['Normal']  # Default style
                
                # Try to match the style if possible
                try:
                    if paragraph.style and paragraph.style.name in target_doc.styles:
                        new_paragraph.style = target_doc.styles[paragraph.style.name]
                except:
                    pass
                
                # Copy run formatting
                for i, run in enumerate(paragraph.runs):
                    if i < len(new_paragraph.runs):
                        new_run = new_paragraph.runs[i]
                        # Copy basic formatting
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.underline = run.underline
                        # Font size if specified
                        if run.font.size:
                            new_run.font.size = run.font.size
            
            # Copy all tables
            for table in source_doc.tables:
                copy_table(table, target_doc)
        
        # Save the merged document
        target_doc.save(target_filename)
        return f"Successfully merged {len(source_filenames)} documents into {target_filename}"
    except Exception as e:
        return f"Failed to merge documents: {str(e)}"
