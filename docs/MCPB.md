# Deployment with MCPB for Claude Desktop

Deployment using **MCPB** packages (`.mcpb`) is the recommended and easiest method to use, specially optimized for **Claude Desktop**. MCPB simplifies the packaging process, eliminating the need to configure complex dependencies and enabling one-click integration.

[Official MCPB repository and documentation](https://github.com/modelcontextprotocol/mcpb)

---

## Benefits of using MCPB with Claude Desktop

- **"Plug-and-Play" Installation**: No need to manually edit the `claude_desktop_config.json` file.
- **Visual Configuration**: Claude Desktop will read the variables exposed by the package and integrate a visual form where you simply type your values.
- **Total Portability**: The packaging contains everything you need and the pre-configured `manifest.json` at the root of the project ensures a stable integration.

---

## Installation and Usage Steps

### 1. Prepare the environment
Make sure you have the **MCPB** CLI installed. You can refer to the official GitHub repository for different installation methods on your system.

### 2. Generate the Pack (Package)
Open a terminal at the root of this project (`mcp_word_office`) and run the packaging command:

```bash
mcpb pack
```

This will build your server based on the `manifest.json` file and generate a compatible package (a `.mcpb` file).

### 3. Install in Claude Desktop
Once the package is generated, install it in your client:

1. Open the **Claude Desktop** application.
2. During the installation process of this package (by dragging or importing it), the Claude graphical interface will guide you.
3. **Configure Permissions**: Claude Desktop will detect the variable defined in the manifest and will prompt you to enter its value on screen. You must configure:
   - `MCP_ALLOWED_DIRECTORIES`: Type the absolute path to your files or the Documents folder that you want Claude to read and edit (e.g., `C:\Users\yourUser\Documents`).

### 4. Restart and Start (Optional)
Depending on your version of Claude Desktop, you might be prompted to restart the application. Once it's loaded, interact with it directly in the chat asking to manipulate a *Word* file!

---

## Included `manifest.json` file

This project already comes with a `manifest.json` file configured at the root to facilitate out-of-the-box support with MCPB. It has been defined so that:
- It starts using the best execution path in your environment.
- It explicitly and securely requires the definition of the `MCP_ALLOWED_DIRECTORIES` variable.

If any issues arise or you want more advanced execution parameters for packaging, check the [Official MCPB GitHub](https://github.com/modelcontextprotocol/mcpb).
