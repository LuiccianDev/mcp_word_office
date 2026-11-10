# Deployment with MCPB for Claude Desktop

This guide explains how to bundle and install **Word MCP Server** as an MCPB package (`.mcpb`) for Claude Desktop.

---

## 1. Overview

MCPB bundles the server manifest (`manifest.json`) and repository resources into a portable package file for Claude Desktop. It enables visual configuration in Claude Desktop's settings UI while executing the server with `uv`.

---

## 2. Prerequisites

Before creating or installing an MCPB package, verify that your environment has:

- **Astral `uv`**: Installed and accessible on your system `PATH` (e.g. `uv --version`).
- **Python 3.11+**: Supported runtime for Word MCP.
- **Claude Desktop**: Installed on macOS or Windows.
- **MCPB CLI**: Installed according to the official MCPB documentation.

---

## 3. Build

From the root of the repository, build the package using the MCPB CLI:

```bash
mcpb pack
```

This reads `manifest.json` from the current directory and creates a `.mcpb` bundle containing the package metadata, configuration options, and tool definitions.

---

## 4. Install

1. Open **Claude Desktop**.
2. Drag and drop the generated `.mcpb` file into Claude Desktop, or use **Settings > Extensions > Install Extension**.
3. Claude Desktop installs the bundle and prompts for required configuration parameters.

---

## 5. Configuration

The extension exposes a directory configuration field mapped from the manifest:

- **Directory (`MCP_ALLOWED_DIRECTORIES`)**: Enter the absolute path to the directory containing your Word documents (e.g., `/Users/username/Documents` or `C:\Users\username\Documents`).
- All document read, write, and batch operations will be restricted strictly to this designated path and its subdirectories.

---

## 6. Verification

Once installed:

1. Open a new conversation in Claude Desktop.
2. Confirm the extension indicator shows **word_mcp** as connected.
3. Verify that `build_document` and other tools appear in the available tool list.
4. Issue a test prompt:
   ```text
   Create a document named test.docx with a heading "Test Document" and a short paragraph.
   ```

---

## 7. Troubleshooting

- **`uv: command not found`**: Ensure Astral `uv` is installed and that its directory is in your user `PATH`. Restart Claude Desktop after updating `PATH`.
- **Invalid Directory / Permission Denied**: Check that the path provided in the directory configuration exists, is absolute, and grants read/write permissions to your user account.
- **Python Version Mismatch**: Word MCP requires Python 3.11 or newer. Run `uv python install 3.11` to allow `uv` to provision a compatible runtime automatically.

---

## 8. Reference

- [Official MCPB Repository](https://github.com/modelcontextprotocol/mcpb)
- [Word MCP Tools Reference](TOOLS.md)
