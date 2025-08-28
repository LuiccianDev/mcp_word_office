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

> 💡 **Note**: This project is designed to be used with MCP-compatible clients such as Claude, allowing Word document manipulation via natural language instructions.

## 📑 Table of Contents

- [🖥️ System Requirements](#️-system-requirements)
- [⚙️ Installation](#️-installation)
- [🗂️ Project Structure](#️-project-structure)
- [🔧 Tool API](#-tool-api)
- [🔒 Security](#-security)
- [🤝 Contribution](#-contribution)
- [📜 License](#-license)

## 🖥️ System Requirements

### 📋 Minimum Requirements

- **Python**: 3.13 or higher
- **UV Package Manager**: [Install UV](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or use pip
- **Git**: For cloning the repository
- **Desktop Extensions (DXT)**: for creating .dxt packages for Claude desktop [Install DXT](https://github.com/anthropics/dxt)

## ⚙️ Installation

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

> ℹ️ The above command will install both main and development dependencies.

## 🔌 MCP Client Configuration

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

### DXT Package Deployment

**Best for**: Integrated DXT ecosystem users who want seamless configuration management.

1. **Package the project**:

   ```bash
   dxt pack
   ```

2. **Configuration**: The DXT package automatically handles dependencies and provides user-friendly configuration through the manifest.json:
   - `MCP_ALLOWED_DIRECTORIES`: Base directory for file operations

3. **Usage**: Once packaged, the tool integrates directly with DXT-compatible clients with automatic user configuration variable substitution.

4. **Server Configuration**: This project includes the [manifest.json](manifest.json) file for building the .dxt package.

For more details see [DXT Package Documentation](https://github.com/anthropics/dxt).

### 🔧 Key Environment Variables

| Variable                  | Description                                              | Example                                 |
| ------------------------- | -------------------------------------------------------- | --------------------------------------- |
| `MCP_ALLOWED_DIRECTORIES` | Directories accessible by the server (comma separated)   | `"\Users\User\Documents,.Projects"`     |

## 📦 Docker Support

Para instrucciones detalladas sobre cómo construir y ejecutar el servidor MCP Word Office en Docker, consulta la siguiente guía:

- [Guía Docker: Docker.md](./Docker.md)

## 🔒 Security Considerations

- 🔐 **Allowed Directories**: Limit accessible directories to only those necessary.
- 🛡️ **Virtual Environment**: Always use a virtual environment to isolate dependencies.
- 🔄 **Updates**: Keep the server updated with the latest security patches.
- 👥 **Permissions**: Ensure file permissions are properly set.

## 🗂️ Project Structure

```text
src/
│  └──📁 word_mcp/            # Main server package
│      ├── 📁 core/           # Main Word manipulation logic
│      ├── 📁 tools/          # Exposed MCP tools
│      ├── 📁 utils/          # Utilities and helper functions
│      ├── 📁 prompts/        # Prompt templates for MCP
│      ├── 📁 validation/     # Input validation
│      └── main.py            # Main entry point
│
├── 📁 tests/                  # Unit tests
├── 📄 README.md               # This file
└── 📄 pyproject.toml          # Project configuration
```

### 📋 Directory Description

- **`word_mcp/`**: Contains all server source code.  
  - **`core/`**: Central logic for Word document manipulation.
  - **`tools/`**: Implementation of tools exposed via MCP.
  - **`utils/`**: Shared helper functions.
  - **`prompts/`**: Templates for generating instructions for the language model.
  - **`validation/`**: Input and parameter validation.

- **`tests/`**: Unit and integration tests to ensure proper functionality.

## 🔧 Tool API

The MCP Word server exposes a comprehensive set of tools organized into logical categories for easy Word document manipulation.

> ℹ️ For detailed documentation of each tool, see [TOOLS.md](TOOLS.md).

## 🔒 Security

### Security Considerations

1. **Input Validation**

   - All functions perform strict parameter validation
   - Strongly typed data types are used
   - File path sanitization is applied

2. **File Security**

   - Use of `MCP_ALLOWED_DIRECTORIES` to restrict access
   - Secure handling of temporary files
   - MIME type validation for uploaded files

3. **Best Practices**
   - Code reviewed with `mypy` for type safety
   - Static analysis with `ruff`
   - Unit tests for security cases

## 🤝 Contribution

Contributions are welcome. Please read the contribution guidelines before submitting pull requests.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

<div align="center">
  <p><strong>MCP Word Office Server</strong></p>
  <p>Empowering AI assistants with comprehensive Word manipulation capabilities</p>
  <p>
    <a href="https://github.com/LuiccianDev/mcp_word_office">🏠 GitHub</a> •
    <a href="https://modelcontextprotocol.io">🔗 MCP Protocol</a> •
    <a href="https://github.com/LuiccianDev/mcp_word_office/blob/main/TOOLS.md">📚 Tool Documentation</a>
  </p>
  <p><em>Created with by LuiccianDev</em></p>
</div>
