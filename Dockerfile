FROM python:3.13-slim

# System dependency for headless PDF conversion (LibreOffice).
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

# Astral uv, used for reproducible dependency installation.
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

WORKDIR /app

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    PATH="/app/.venv/bin:$PATH" \
    MCP_ALLOWED_DIRECTORIES=/documents

# Install only runtime dependencies from the lockfile.
COPY pyproject.toml uv.lock README.md LICENSE ./
COPY src ./src
RUN uv sync --frozen --no-dev

RUN mkdir -p /documents

CMD ["luiccian-word-mcp"]
