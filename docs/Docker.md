# Docker Deployment

This guide builds the Word MCP server as a Docker image and connects it to an MCP client over `stdio`.

The image bundles LibreOffice Writer, which the server uses for headless PDF conversion on Linux.

---

## 1. Build the Docker image

Open a terminal at the repository root and run:

```bash
docker build -t word-mcp .
```

---

## 2. Run the server

```bash
docker run --rm -i \
  -v /path/to/your/documents:/documents \
  -e MCP_ALLOWED_DIRECTORIES=/documents \
  word-mcp
```

- `-i` keeps standard input open for the `stdio` transport.
- `-v` mounts the directory that contains the Word documents. Replace `/path/to/your/documents` with the absolute path on your host.
- `-e MCP_ALLOWED_DIRECTORIES` restricts file access to the mounted directory.

---

## 3. Connect an MCP client

### Claude Desktop

Add the following entry to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "docker",
      "args": [
        "run", "--rm", "-i",
        "-v", "/path/to/your/documents:/documents",
        "--name", "word-mcp",
        "-e", "MCP_ALLOWED_DIRECTORIES=/documents",
        "word-mcp"
      ]
    }
  }
}
```

Adjust the mounted path to match your system.

---

## 4. Usage flow

1. Start the MCP client.
2. The client launches the container with the configuration above.
3. All communication uses `stdio`; no ports need to be exposed.
4. Ask the client to create, edit, or convert Word documents inside the mounted directory.

---

## 5. Security recommendations

- Always set `MCP_ALLOWED_DIRECTORIES` to restrict access to the required folder only.
- Mount only the volumes that are strictly necessary.
- Review the permissions of the shared files.
- Do not open firewall ports for the container.

---

## 6. Resources

- [Model Context Protocol](https://modelcontextprotocol.io)
- [Claude Desktop](https://www.anthropic.com/)
- [Word MCP repository](https://github.com/LuiccianDev/word-mcp)
