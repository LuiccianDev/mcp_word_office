# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

MCP Office Word Server is a Python-based Model Context Protocol (MCP) server that provides comprehensive Microsoft Word (.docx) document manipulation capabilities. It's designed to work with MCP-compatible clients like Claude to enable programmatic document processing through natural language instructions.

## Architecture

The project follows a modular, layered architecture:

- **`src/word_mcp/main.py`**: FastMCP server initialization and entry point
- **`src/word_mcp/tools/`**: MCP tool implementations organized by functionality
- **`src/word_mcp/core/`**: Core document processing logic and utilities
- **`src/word_mcp/prompts/`**: MCP prompt templates and registration
- **`src/mcp_server.py`**: Run script for starting the server

### Tool Categories

Tools are organized into logical groups:

- **Document Tools**: Create, copy, merge, list documents (`document_tools.py`)
- **Content Tools**: Add paragraphs, headings, tables, images (`content_tools.py`)
- **Format Tools**: Text formatting, styling, table formatting (`format_tools.py`)
- **Protection Tools**: Document protection and unprotection (`protection_tools.py`)
- **Footnote Tools**: Footnote and endnote management (`footnote_tools.py`)
- **Extended Tools**: Advanced operations like PDF conversion (`extended_document_tools.py`)

All tools are registered via `tools/register_tools.py` using the FastMCP decorator pattern.

## Development Commands

### Environment Setup

```bash
# Create and activate virtual environment (using uv - recommended)
uv venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# Install dependencies (development mode)
uv pip install -e ".[dev]"
```

### Code Quality & Formatting

```bash
# Format code with Black
black src/

# Sort imports with isort
isort src/

# Lint with Ruff (with auto-fix)
ruff check --fix src/

# Type checking with MyPy
mypy src/

# Run all pre-commit hooks
pre-commit run --all-files
```

### Testing

```bash
# Run all tests
pytest

# Run specific test file
pytest tests/test_document_tools.py

# Run with coverage
pytest --cov=src/word_mcp

# Run async tests (most tools are async)
pytest -k "asyncio"
```

### Running the Server

```bash
# Development mode (from project root)
python src/mcp_server.py

# Set allowed directories for security
set MCP_ALLOWED_DIRECTORIES=C:\Users\Documents,C:\Projects
python src/mcp_server.py
```

## Development Guidelines

### Code Standards

- Python 3.13+ required
- **Naming**: snake_case for functions/variables, PascalCase for classes, UPPER_SNAKE_CASE for constants
- **Type Hints**: Mandatory throughout codebase, avoid `Any`
- **Error Handling**: Use custom exceptions from `core/exceptions.py` (DocumentProcessingError, ValidationError, etc.)
- **Input Validation**: Use Pydantic for input validation (as per project rules)

### Testing Patterns

- All tools are async and require `@pytest.mark.asyncio`
- Use `tmp_path` fixture for file operations
- Set `MCP_ALLOWED_DIRECTORIES` environment variable in tests
- Test fixtures create temporary .docx files for testing

### Security Considerations

- **Directory Access**: Server respects `MCP_ALLOWED_DIRECTORIES` environment variable
- **File Validation**: MIME type validation for uploaded files
- **Path Sanitization**: All file paths are validated and sanitized

## Key Dependencies

- **python-docx**: Core Word document manipulation
- **mcp[cli]**: Model Context Protocol implementation
- **msoffcrypto-tool**: Document protection/encryption
- **docx2pdf**: PDF conversion capabilities
- **FastMCP**: Simplified MCP server framework

## Configuration

### Environment Variables

- `MCP_ALLOWED_DIRECTORIES`: Comma-separated list of accessible directories (security)

### MCP Client Configuration

```json
{
  "mcp-word-office": {
    "command": "python",
    "args": ["path/to/mcp_server.py"],
    "env": {
      "MCP_ALLOWED_DIRECTORIES": "path/to/documents"
    }
  }
}
```

## Tool Development Pattern

New tools should follow this pattern:

1. **Location**: Add to appropriate `tools/` module
2. **Async**: All tools must be async functions
3. **Error Handling**: Use custom exceptions from `core/exceptions.py`
4. **Validation**: Validate inputs, especially file paths
5. **Registration**: Register in `tools/register_tools.py`
6. **Testing**: Add comprehensive async tests
7. **Documentation**: Update TOOLS.md with new tool details

Example tool structure:

```python
async def new_tool(filename: str, param: str) -> dict[str, str]:
    """Tool description with parameters and return type."""
    try:
        # Validate inputs
        # Process document
        # Return structured result
        return {"status": "success", "message": "Operation completed"}
    except DocumentProcessingError as e:
        return {"status": "error", "message": str(e)}
```
