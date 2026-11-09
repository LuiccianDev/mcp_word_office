<div align="center">
  <h1>Word MCP Server</h1>
  <p>
    <em>High-performance Model Context Protocol (MCP) server for programmatic Word (.docx) document generation and manipulation</em>
  </p>

  [![PyPI Version](https://img.shields.io/pypi/v/luiccian-word-mcp.svg)](https://pypi.org/project/luiccian-word-mcp/)
  [![Python Version](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/)
  [![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)
  [![CodeQL](https://github.com/LuiccianDev/word-mcp/actions/workflows/codeql.yml/badge.svg)](https://github.com/LuiccianDev/word-mcp/actions/workflows/codeql.yml)
  [![Security Policy](https://img.shields.io/badge/security-policy-blue.svg)](SECURITY.md)
  [![Rules for AI](https://img.shields.io/badge/rules-AGENTS.md-blue.svg)](AGENTS.md)
  [![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
  [![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)

</div>

**Word MCP Server** (`word-mcp`, PyPI: [`luiccian-word-mcp`](https://pypi.org/project/luiccian-word-mcp/)) is a high-performance Python server implementing the Model Context Protocol (MCP) to provide advanced, token-optimized Word document (`.docx`) creation, batch synthesis, and inspection capabilities for AI assistants (Claude, Cursor, Copilot, Cline).

Designed from the ground up for maximum token efficiency, it replaces chatty micro-iterations with single-pass document construction—saving over 90% of model tokens and conversational turns.

---

## Key Architectural Highlights

- **Single-Pass Document Builder**: Create complete, publication-ready documents (headers, footers, page numbers, margins, headings, paragraphs with markdown styling, tables, lists) in **one single MCP call** (`word_build_document`).
- **Markdown Direct Generation**: Instantly turn structured markdown text into professionally styled Word documents (`word_create_from_markdown`).
- **Visual Design Presets**: Built-in typography and spacing standards (`modern`, `formal`, `academic`, `compact`).
- **Lightweight Outline Inspection**: Extract heading hierarchies and document stats without dumping full body text (`word_get_document_outline`), saving context tokens.
- **Batch Content Ingestion**: Append multiple sections, tables, and lists to existing documents in a single atomic operation (`word_batch_add_content`).
- **Protection & Security**: Password protection, editing restrictions, file integrity checks, and headless PDF conversion.
- **Zero-Install Execution**: Instant launch with `uvx luiccian-word-mcp`.

---

## Quick Start with `uvx` (Recommended)

No manual clone or installation required. You can run the server directly using [`uv`](https://docs.astral.sh/uv/):

```bash
# Run with stdio transport (default for AI IDEs)
uvx luiccian-word-mcp

# Or run with SSE transport
uvx luiccian-word-mcp --transport sse --host 127.0.0.1 --port 8000
```

---

## MCP Client Configuration

### 1. Claude Desktop

Add this configuration to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "MCP_ALLOWED_DIRECTORIES": "/path/to/your/documents"
      }
    }
  }
}
```

> **Tip (Local checkout)**: If running from a local clone:
> ```json
> {
>   "mcpServers": {
>     "word-mcp": {
>       "command": "uv",
>     "args": ["run", "--directory", "/path/to/word-mcp", "luiccian-word-mcp"],
>       "env": {
>         "MCP_ALLOWED_DIRECTORIES": "/path/to/your/documents"
>       }
>     }
>   }
> }
> ```

### 2. Cursor (.cursor/mcp.json)

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "uvx",
      "args": ["luiccian-word-mcp"],
      "env": {
        "MCP_ALLOWED_DIRECTORIES": "${workspaceFolder}"
      }
    }
  }
}
```

---

## Local Development with `uv`

```bash
# 1. Clone repository
git clone https://github.com/LuiccianDev/word-mcp.git
cd word-mcp

# 2. Sync virtual environment and dependencies
uv sync --all-groups

# 3. Run server
uv run luiccian-word-mcp

# 4. Run tests and linter
uv run pytest
uv run ruff check .
```

---

## CI/CD & PyPI Trusted Publishing

The repository includes automated GitHub Actions workflows powered by Astral `uv`:

- **CI (`.github/workflows/ci.yml`)**: Runs tests & lint across Python 3.11, 3.12, and 3.13 on Ubuntu. Uses `paths-ignore` and concurrency cancellation to preserve GitHub Actions minutes.
- **CodeQL Security (`.github/workflows/codeql.yml`)**: Continuous automated security analysis and vulnerability scanning for Python code with GitHub Advanced Security.
- **Automated Version Bump & PyPI Release (`.github/workflows/bump-and-release.yml`)**: 1-click semantic version bump (`patch`, `minor`, `major`). Automatically updates `pyproject.toml`, promotes `CHANGELOG.md`, creates git commit & tag `v*`, builds with `uv build`, attests cryptographic provenance (PEP 740), generates GitHub Release, and publishes to PyPI.
- **Releases (`.github/workflows/release.yml`)**: Triggered on `v*` tags or manually via workflow dispatch. Builds wheel & sdist with `uv build`, auto-extracts release notes from `CHANGELOG.md`, and creates a GitHub Release.
- **PyPI Publishing (`.github/workflows/publish-pypi.yml`)**: Automated PyPI deployment using **Trusted Publishing (OIDC)** and PEP 740 cryptographic build provenance attestation (`actions/attest-build-provenance@v2`).

### PyPI Trusted Publisher Configuration

To set up or verify Trusted Publishing on [PyPI.org](https://pypi.org/manage/account/publishing/):

| Setting | Value |
| :--- | :--- |
| **PyPI Project Name** | `luiccian-word-mcp` |
| **Owner** | `LuiccianDev` |
| **Repository name** | `word-mcp` |
| **Workflow name** | `publish-pypi.yml` |
| **Environment name** | `pypi` |

---

## Token Optimization: Single-Pass vs. Legacy

| Metric | Legacy Micro-Calls | Word MCP Single-Pass |
| :--- | :--- | :--- |
| **Tool Calls** | 15 – 35 calls per document | **1 single call** (`word_build_document`) |
| **Token Usage** | ~15,000 – 30,000 tokens | **~800 – 1,500 tokens** (90%+ savings) |
| **Disk Operations** | Multiple file writes | Single in-memory assembly & write |
| **Generation Time** | 30 – 60 seconds | **< 0.25 seconds** |

---

## Visual Design Presets

| Preset | Font Family | Body Size | Line Spacing | Margins | Best For |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `modern` (Default) | Calibri | 11.0 pt | 1.15 | 2.5 cm | Executive summaries, business proposals, corporate reports |
| `formal` | Times New Roman | 12.0 pt | 1.50 | Top/Bottom 2.5cm, Left 3.0cm | Legal contracts, official letters, meeting minutes |
| `academic` | Times New Roman | 12.0 pt | 2.00 (Double) | 2.54 cm (1 inch) | University papers, research articles, theses (APA/MLA) |
| `compact` | Arial | 10.0 pt | 1.00 (Single) | 2.0 cm | Invoices, datasheets, internal memos |

---

## Modular Skills Ecosystem

The repository provides 3 modular, specialized skills in `skills/` for LLMs:

- **`word-document-creator`**: Instructs the model to build complete documents in 1 single pass with headers, footers, pagination, tables, and presets.
- **`word-document-editor`**: Focuses on lightweight outline inspection, batch mutations, search & replace, and document merging.
- **`word-document-security`**: Handles password locking, editing restrictions, and PDF export for distribution.

---

## Project Structure

```text
src/
└── word_mcp/
    ├── core/               # Word document manipulation engines
    │   ├── batch.py        # High-performance single-pass builder & markdown parser
    │   ├── document.py     # Document lifecycle & outline extraction
    │   ├── content.py      # Paragraphs, headings, and lists
    │   ├── tables.py       # Table styling & cell manipulation
    │   └── protection.py   # Password hashing & editing locks
    ├── tools/              # Exposed MCP tools (consolidated)
    │   ├── batch_tools.py  # word_build_document, word_create_from_markdown
    │   ├── document_tools.py # Lifecycle, text, outline, info, PDF
    │   ├── content_tools.py  # Paragraphs, headings, links, bookmarks
    │   ├── format_tools.py   # Styles, headers, footers, page layout
    │   └── protection_tools.py # Protect, unprotect, verify
    ├── models/             # Pydantic response models
    ├── validation/         # Input validators & parameter alias tolerance
    ├── server.py           # FastMCP server registration
    └── __main__.py         # Server entry point
```

---

## Package Distribution & Repository Scope

The distribution boundary between the PyPI package and the repository is strictly defined:

- **PyPI Package (`luiccian-word-mcp`)**: Bundles **exclusively** the runtime code located in [`src/word_mcp/`](file://src/word_mcp/), accompanied by `README.md`, `LICENSE`, and `pyproject.toml`.
- **Repository Scope**: Test suites (`tests/`), automation scripts (`scripts/`), agent skills (`skills/`), documentation manuals (`docs/`), and GitHub Actions workflows (`.github/`) belong solely to repository infrastructure and are explicitly excluded from the PyPI wheel and sdist distributions via Hatchling build targets.

> [!NOTE]
> Installing via `uvx luiccian-word-mcp` or `pip install luiccian-word-mcp` downloads only the core library and runtime dependencies, ensuring zero overhead and rapid startup.

---

## Security & Guidelines

- **Security Policy**: See [`SECURITY.md`](SECURITY.md) for vulnerability disclosure, OOXML hash verification, and threat modeling.
- **Rules for AI Assistants**: See [`AGENTS.md`](AGENTS.md) for official development directives, coding standards, and token optimization rules.
- **Directory Isolation**: Set `MCP_ALLOWED_DIRECTORIES` to restrict document operations strictly to designated folders. Path canonicalization prevents directory traversal.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

<div align="center">
  <p><strong>Word MCP Server</strong> • Maintained by <a href="https://github.com/LuiccianDev">LuiccianDev</a></p>
</div>
