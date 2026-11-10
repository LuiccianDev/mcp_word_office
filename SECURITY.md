# Security Policy for Word MCP (`word-mcp`)

Document integrity, sandbox isolation, and host security are fundamental priorities for **Word MCP Server** (`word-mcp`, PyPI: [`luiccian-word-mcp`](https://pypi.org/project/luiccian-word-mcp/)).

Because this Model Context Protocol server interacts with local filesystems, generates documents, and integrates with LLM agents (Claude, Cursor, Cline, Copilot), strict security controls and defense-in-depth principles are enforced across all modules.

---

## Supported Versions

Only the latest major version branch receives active security updates and vulnerability patches.

| Version | Supported | Notes |
| :--- | :--- | :--- |
| **3.x.x** | :white_check_mark: Yes | Current major branch after the 3.0.0 release |
| **2.x.x** | :x: No | Superseded by 3.0.0; prefixed tool aliases were removed |
| **1.x.x** | :x: No | Deprecated legacy `mcp_word` release |

---

## Security Architecture & Defenses

### 1. Directory Sandboxing (`MCP_ALLOWED_DIRECTORIES`)
By default or via configuration, Word MCP limits document reading, writing, and conversions strictly to explicit allowed directories:
- The server canonicalizes all filesystem paths via `os.path.realpath` and `pathlib.Path.resolve()`.
- Directory traversal attacks using relative dot-segments (`../`, `..\\`) are detected and blocked with `SecurityError` before any file handle is opened.

### 2. Password Protection and Protection Metadata
- `protect_document` encrypts the OOXML container with `msoffcrypto-tool`. The password is not stored in the document or in this repository.
- `unprotect_document` decrypts the container and returns `Incorrect password` when the key is invalid.
- `add_restricted_editing` and `add_digital_signature` write a sidecar JSON file named `<document>.protection` next to the `.docx`. The sidecar stores a salted PBKDF2-HMAC-SHA-256 hash of the password for restricted-editing verification using a unique cryptographic salt per document and constant-time comparison. The `.protection` file is advisory metadata for MCP clients and does not replace native OOXML encryption (`protect_document`).
- `add_digital_signature` writes a visible signature block and stores a SHA-256 hash of the paragraph text. It detects text changes after signing. It is **not a cryptographic signature** and does not verify signer identity.
- `verify_document` reads the sidecar metadata only. It does not validate OOXML package signatures.

### 3. Safe Subprocess Execution
- Subprocess invocations (such as LibreOffice headless PDF conversion) use argument lists directly rather than `shell=True` to eliminate shell injection vulnerabilities.
- Command existence and execution timeouts are enforced to prevent denial of service (DoS).

### 4. Known Limitations
- The `.protection` sidecar is advisory metadata. Microsoft Word does not enforce the `editable_sections` list.
- Page count is not computed by `get_document_info`; rendering is required for an authoritative value.
- PDF conversion depends on LibreOffice on Linux and `docx2pdf` on Windows/macOS. Failure is reported as a structured error.

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
