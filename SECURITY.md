# Security Policy for Word MCP (`word-mcp`)

Document integrity, sandbox isolation, and host security are fundamental priorities for **Word MCP Server** (`word-mcp`, PyPI: [`luiccian-word-mcp`](https://pypi.org/project/luiccian-word-mcp/)).

Because this Model Context Protocol server interacts with local filesystems, generates documents, and integrates with LLM agents (Claude, Cursor, Cline, Copilot), strict security controls and defense-in-depth principles are enforced across all modules.

---

## Supported Versions

Only the latest major version branch receives active security updates and vulnerability patches.

| Version | Supported | Notes |
| :--- | :--- | :--- |
| **2.x.x** | :white_check_mark: Yes | Current production release branch (`word-mcp`) |
| **1.x.x** | :x: No | Deprecated legacy `mcp_word` release |

---

## Security Architecture & Defenses

### 1. Directory Sandboxing (`MCP_ALLOWED_DIRECTORIES`)
By default or via configuration, Word MCP limits document reading, writing, and conversions strictly to explicit allowed directories:
- The server canonicalizes all filesystem paths via `os.path.realpath` and `pathlib.Path.resolve()`.
- Directory traversal attacks using relative dot-segments (`../`, `..\\`) are detected and blocked with `SecurityError` before any file handle is opened.

### 2. Password Hashing & OOXML Encryption
- Document protection operations (`protect_document`, `verify_protection`) implement standard OOXML hash algorithms (SHA-512 with 100,000 iterations) according to ECMA-376 specifications.
- Passwords are never logged in plain text or exposed in response models.

### 3. Safe Subprocess Execution
- Subprocess invocations (such as LibreOffice headless PDF conversion) use argument lists directly rather than `shell=True` to eliminate shell injection vulnerabilities.
- Command existence and execution timeouts are enforced to prevent denial of service (DoS).

---

## Reporting a Vulnerability

If you discover a security vulnerability in **Word MCP Server**, please report it responsibly:

1. **Private Disclosure**: Do **not** open a public issue or discussion on GitHub.
2. **Contact**: Send an email directly to the maintainer at [luicciandev@gmail.com](mailto:luicciandev@gmail.com) with the subject:
   `[SECURITY VULNERABILITY] Word MCP - <Brief Description>`
3. **Information to Include**:
   - Detailed description of the vulnerability and its potential impact.
   - Proof-of-concept (PoC) code or exact steps to reproduce.
   - Affected version(s) and operating system.
   - Any suggested mitigations or patches if available.

### Response Timeline
- **Initial Acknowledgment**: Within 48 hours.
- **Triage & Assessment**: Within 5 business days.
- **Fix & Advisory**: A patched release will be published via GitHub Releases and PyPI with coordinated disclosure.
