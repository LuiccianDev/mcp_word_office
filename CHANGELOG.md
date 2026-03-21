# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Added
- **MCP Document Workflow Skill**: Comprehensive skill for creating and improving high-quality .docx files through a 5-step workflow (brief capture, structure planning, iterative content build, formatting pass, final QA)
- **MCP Word Document Workflow Skill Templates**: Templates for common document types including acta, carta, informe-universitario, propuesta, reporte, solicitud-formal, and syllabus
- **Comprehensive Skill Documentation**: Added detailed SKILL.md, README.md, and reference documentation for the MCP Word Document Workflow
- **PROJECT_SPECIFIC.mcpbignore patterns**: Improved project-specific exclusions for better build optimization
- **Project License File**: Added comprehensive LICENSE file for project licensing information

### Changed
- **FastMCP3 Integration**: Upgraded to FastMCP3 with enhanced Pydantic v2 support
- **Pydantic v2 Migration**: Updated all Pydantic models and validators for Pydantic v2 compatibility
- **Server Implementation**: Refactored `server.py` for better modularity and enhanced tool registration

### Fixed
- Improved handling of file exclusions in `.mcpbignore` configuration

### Repository Structure
- Enhanced documentation and templates in `skills/mcp-word-document-workflow/`
- Updated `.mcpbignore` to include `docs/` and `skills/` directories

---

## [1.1.2] - 2026-03-21

### Added
- **Document Validation Framework**: Comprehensive document validation decorators to verify DOCX file existence, type, and writability
- **Extended Word Document Tools**: Advanced document operations including pagination, complex formatting, and bulk operations
- **Comprehensive Tool Suite**: Document manipulation tools, extended tools, footnote tools, format tools, header/footer tools, link tools, property tools, protection tools with full error handling
- **Server Infrastructure**: Complete MCP server setup with tool registration and context management
- **Response Models**: Structured response models for all tool outputs ensuring consistent client integration

### Changed
- **Project Structure**: Reorganized core modules for better separation of concerns
- **Testing Infrastructure**: Enhanced test suite covering all tool categories and core functionality
- **Manifest Configuration**: Updated MCPBuilder manifest version and configuration

### Fixed
- Document context management and error handling improvements
- Tool response formatting for better client compatibility

---

## [1.1.1] - 2026-03-20

### Added
- **Enhanced Documentation**: Added comprehensive TOOLS.md documentation for all available tools and capabilities

### Changed
- **Docker Support**: Improved Docker configuration and documentation for containerized deployments
- **Dependency Management**: Updated dependencies for better compatibility and performance

### Fixed
- Various bug fixes in document manipulation logic
- Improved error messages for better debugging

---

## [1.1.0] - 2026-03-19

### Added
- **Core Word Document Operations**: Implemented comprehensive document management and manipulation capabilities
- **Style Management**: Enhanced style handling with support for custom styles and formatting
- **Testing Framework**: Added comprehensive test suite for all modules
- **Validation Decorators**: New validation system for document operations
- **Documentation**: Extensive inline documentation and type hints throughout codebase

### Changed
- **Import Organization**: Refactored and cleaned up import statements across all modules for better code organization
- **Module Structure**: Reorganized package structure for improved maintainability
- **Package Version**: Updated from 1.0.1 to 1.1.0

### Fixed
- Import conflicts and circular dependencies
- Module initialization issues

---

## [1.0.1] - 2026-03-15

### Added
- **Docker Support**: Added Dockerfile and .dockerignore for containerized deployments
- **Docker Documentation**: Comprehensive Docker setup and deployment documentation in TOOLS.md

### Changed
- **Package Rename**: Renamed package from `mcp-office-word` to `mcp-word` for better naming consistency
- **Script Updates**: Updated all script and entry-point names to reflect new package name
- **Documentation**: Updated README.md to support additional MCP-compatible clients and improved clarity
- **Manifest Configuration**: Updated manifest.json and configuration files for new package name

### Fixed
- Corrected version numbers in pyproject.toml and __init__.py to 1.0.1
- Fixed trailing spaces and formatting issues in documentation
- Corrected IDE-specific files in .gitignore (.kiro/, .windsurf/)

### Security
- Added proper .gitignore patterns for sensitive and build-related files

---

## [1.0.0] - 2026-02-15

### Added
- **Initial Release**: Complete MCP server implementation for Word document manipulation
- **Document Operations**: Create, read, copy, merge, and convert documents to PDF
- **Content Management**: Add and manipulate headings, paragraphs, tables, pictures, page breaks
- **Search & Replace**: Find and replace content within documents
- **Text Formatting**: Apply comprehensive text formatting including bold, italic, underline, colors
- **Custom Styles**: Create and apply custom document styles
- **Footnotes & Endnotes**: Add, manage, and convert between footnotes and endnotes
- **Security Features**: Password protection, restricted editing mode, digital signature support
- **Headers & Footers**: Create and manage document headers and footers for specific sections
- **Links & Bookmarks**: Insert hyperlinks and internal bookmarks for navigation
- **Document Properties**: Manage core metadata (author, title) and section layouts
- **Python SDK**: Full Python-based MCP server implementation
- **Docker Support**: Dockerfile and documentation for containerized deployment
- **Comprehensive Testing**: Unit tests for core functionality and utilities
- **Documentation**: Complete README with features, installation, and usage instructions

### Dependencies
- `python-docx>=1.0.0`: Core Word document manipulation
- `docx2pdf>=0.1.8`: PDF conversion support
- `msoffcrypto-tool>=5.4.2`: Document protection and encryption
- `fastmcp>=0.9.0`: MCP server framework (FastMCP)

---

## Links

- [GitHub Repository](https://github.com/LuiccianDev/mcp_word_office)
- [GitHub Issues](https://github.com/LuiccianDev/mcp_word_office/issues)
- [MCP Protocol](https://modelcontextprotocol.io)
- [Python-docx Documentation](https://python-docx.readthedocs.io)
