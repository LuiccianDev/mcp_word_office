<div align="center">
  <h1> MCP Office Word Server</h1>
  <p>
    <em>Powerful server for programmatic manipulation of Word documents (.docx) via MCP</em>
  </p>

  [![Python Version](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/)
  [![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)

</div>

**MCP Office Word Server** is a Python server implementing the Model Context Protocol (MCP) to provide advanced capabilities for manipulating Microsoft Word (`.docx`) documents. This server enables automation of complex document processing tasks programmatically.

> **Note**: This project is designed to be used with MCP-compatible clients such as Claude, Cursor or others IDEs, allowing Word document manipulation via natural language instructions.

## Features

- **Document Operations**: Create, read, copy, merge, and convert to PDF.
- **Content Management**: Add headings, paragraphs, tables, pictures, and page breaks. Search, replace, and delete content.
- **Formatting**: Apply text formatting, custom styles, and table formatting.
- **Footnotes & Endnotes**: Add and convert footnotes and endnotes, and customize their styles.
- **Security & Protection**: Add password protection, restricted editing, and digital signatures.
- **Headers & Footers**: Add or update primary headers and footers for specific sections.
- **Links & Bookmarks**: Insert hyperlinks and internal bookmarks for navigation.
- **Document Properties**: Manage core metadata (author, title) and section page layouts (portrait/landscape).

## Table of Contents

- [Features](#features)
- [Project Structure](#project-structure)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [MCP Client Configuration](#mcp-client-configuration)
- [Docker Support](#docker-support)
- [Tool API](#tool-api)
- [Security](#security)
- [Changelog](#changelog)
- [Contribution](#contribution)
- [License](#license)

## Project Structure

```text
src/
│  └── 📁 mcp_word/           # Main server package
│      ├── 📁 core/           # Main Word manipulation logic
│      ├── 📁 tools/          # Exposed MCP tools
│      ├── 📁 exception/      # Centralized exception and unprotect handling
│      ├── 📁 models/         # Response models and schemas
│      ├── 📁 utils/          # Utilities and helper functions
│      ├── 📁 validation/     # Input validation
│      ├── 📄 server.py       # Main entry point
│      └── 📄 __main__.py     # Alternative module entry point
│
├── 📁 tests/                  # Unit tests
├── 📄 README.md               # This file
└── 📄 pyproject.toml          # Project configuration
```

### Directory Description

- **`mcp_word/`**: Contains all server source code.
  - **`core/`**: Central logic for Word document manipulation.
  - **`tools/`**: Implementation of tools exposed via MCP.
  - **`exception/`**: Exception handling and protection-related exception tools.
  - **`models/`**: Shared response models used across MCP tools.
  - **`utils/`**: Shared helper functions.
  - **`validation/`**: Input and parameter validation.
  - **`server.py`**: Main MCP server entry point.
  - **`__main__.py`**: Alternative entry point for module execution.

- **`tests/`**: Unit and integration tests to ensure proper functionality.

## System Requirements

### Minimum Requirements

- **Python**: 3.13 or higher
- **UV Package Manager**: [Install UV](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or use pip
- **Git**: For cloning the repository
- **Desktop Extensions (MCPB)**: for creating .mcpb packages for Claude desktop [Install MCPB](https://github.com/modelcontextprotocol/mcpb)

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/LuiccianDev/mcp_office_word.git
cd mcp_office_word
```

### 2. Set up virtual environment with uv

uv is the recommended package manager for this project. It automatically creates and activates a virtual environment:

```bash
# Create and activate virtual environment
uv venv

# On Windows:
.venv\Scripts\activate

# On macOS/Linux:
source .venv/bin/activate
```

### 3. Install dependencies

Install project dependencies using uv:

```bash
uv pip install -e ".[dev]"
```

> The above command will install both main and development dependencies.

## MCP Client Configuration

### Basic Configuration

To integrate the MCP Office Word server with compatible clients like Claude, follow these steps:

1. **Start the server** following the instructions in the section
2. **Configure your MCP client** with the following parameters:

```bash
# create packages
uv build
#install packages
pip install dist/file*.whl
```

The next steps are configurations in MCP

```json
{
  "mcp-word": {
      "command": "uv",
      "args": ["run", "mcp_word"],
    "env": {
      "MCP_ALLOWED_DIRECTORIES": "Users/path/to/your/documents"
    }
  }
}
```

Or use this configuration (less recommended):

```json
{
   "mcp-word": {
      "command": "/Users/user/to/repo/.venv/Scripts/python",
      "args": [
        "/Users/user/to/repo/src/mcp_word/server.py"
      ],
   "env": {
      "MCP_ALLOWED_DIRECTORIES": "/Users/user/to/WorksPath"
   },
   "type": "stdio"
  },
}
```

### Key Environment Variables

| Variable                  | Description                                              | Example                                 |
| ------------------------- | -------------------------------------------------------- | --------------------------------------- |
| `MCP_ALLOWED_DIRECTORIES` | Directories accessible by the server (comma separated)   | `"\Users\User\Documents,.Projects"`     |

### MCPB Package Deployment

For detailed instructions on creating and using MCPB packages (.mcpb), please see:

- [MCPB Guide: mcpb.md](docs/MCPB.md)

## Docker Support

For detailed instructions on building and running the MCP Word Office server in Docker, please see the following guide:

- [Docker Guide: Docker.md](docs/Docker.md)

## Tool API

The MCP Word server exposes a comprehensive set of tools organized into logical categories for easy Word document manipulation.

> For detailed documentation of each tool, see [TOOLS.md](docs/TOOLS.md).

## Security

### Considerations & Best Practices

1. **Input Validation**
   - All functions perform strict parameter validation
   - Strongly typed data types are used
   - File path sanitization is applied

2. **File Security**
   - Use of `MCP_ALLOWED_DIRECTORIES` to restrict access (limit accessible directories to only those necessary).
   - Secure handling of temporary files
   - MIME type validation for uploaded files
   - Ensure file permissions are properly set.

3. **Execution Environment**
   - Always use a virtual environment to isolate dependencies.
   - Keep the server updated with the latest security patches.
   - Code reviewed with `mypy` for type safety
   - Static analysis with `ruff`
   - Unit tests for security cases

## Contribution

Contributions are welcome. Please read the contribution guidelines before submitting pull requests.

## Changelog

For detailed information about changes, improvements, and bug fixes in each version, please see the [CHANGELOG.md](CHANGELOG.md) file.

Key information:
- **Current Version**: 1.1.2
- **Latest Features**: FastMCP3 integration, Pydantic v2 support, MCP Word Document Workflow Skill
- **Unreleased**: See the [Unreleased] section in CHANGELOG.md for upcoming features

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

<div align="center">
  <p><strong>MCP Word Office Server</strong></p>
  <p>Empowering AI assistants with comprehensive Word manipulation capabilities</p>
  <p>
    <a href="https://github.com/LuiccianDev/mcp_word_office">🏠 GitHub</a> •
    <a href="https://modelcontextprotocol.io">🔗 MCP Protocol</a> •
    <a href="https://github.com/LuiccianDev/mcp_word_office/blob/main/docs/TOOLS.md">📚 Tool Documentation</a>
  </p>
  <p><em>Created with by LuiccianDev</em></p>
</div>
