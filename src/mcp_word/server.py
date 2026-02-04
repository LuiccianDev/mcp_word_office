"""
Main entry point for the Word Document MCP Server.

Acts as the central controller for the MCP server that handles Word document operations.
"""

from mcp.server.fastmcp import FastMCP

from mcp_word.prompts.register_prompts import register_prompts
from mcp_word.tools.register_tools import register_all_tools


def create_server() -> FastMCP:
    """Create and configure the Word Document MCP Server."""

    mcp = FastMCP(
        name="word_mcp",
        instructions="""# MCP Word Server

MCP Word Server provides comprehensive Word document (.docx) manipulation capabilities for AI assistants.

## Capabilities

### Document Operations
- **Create documents**: Generate new Word documents with optional metadata (title, author)
- **Read documents**: Extract text, information, and structure from existing documents
- **Copy documents**: Duplicate documents with automatic naming
- **Merge documents**: Combine multiple documents into one
- **Convert to PDF**: Export documents to PDF format

### Content Management
- **Add content**: Headings, paragraphs, tables, pictures, page breaks
- **Search & replace**: Find and replace text throughout documents
- **Delete content**: Remove paragraphs by index

### Formatting
- **Text formatting**: Bold, italic, underline, colors, font size, font family
- **Custom styles**: Create reusable character and paragraph styles
- **Table formatting**: Borders, shading, header rows

### Footnotes & Endnotes
- **Add footnotes**: Reference notes at page bottom
- **Add endnotes**: Reference notes at document end
- **Convert between**: Transform footnotes to endnotes and vice versa
- **Customize styling**: Configure footnote/endnote appearance

### Document Protection
- **Password protection**: Encrypt documents with passwords
- **Restricted editing**: Limit editing to specific sections
- **Digital signatures**: Sign documents digitally
- **Verification**: Check document protection and signature status

## Usage Patterns

### Automation
- Generate reports from templates
- Batch process multiple documents
- Extract data from existing documents
- Apply consistent formatting across documents

### Document Creation
- Create new documents with structured content
- Add tables with data
- Insert images
- Build tables of contents

## Security

- All file operations are restricted to directories specified in `MCP_ALLOWED_DIRECTORIES`
- Document protection uses strong encryption
- Digital signatures verify document integrity

## Tool Naming

All tools use the `word_` prefix for clarity and to avoid conflicts with other MCP servers. Legacy names without the prefix are also available for backward compatibility.""",
        dependencies=["python-docx"],
        on_duplicate_tools="error",
    )

    register_all_tools(mcp)
    register_prompts(mcp)

    return mcp


if __name__ == "__main__":
    mcp = create_server()
    mcp.run(transport="stdio")
