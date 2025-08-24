# register_prompts.py - registro de prompts
from collections.abc import Awaitable, Callable
from typing import Any

from mcp.server.fastmcp import FastMCP

from mcp_word.prompts.prompts import PROMPT_REGISTRY, PromptTemplate


def register_prompts(mcp: FastMCP) -> None:
    """Register all prompts with the MCP server."""

    for _, prompt_template in PROMPT_REGISTRY.items():
        # Crear función closure para cada prompt
        def create_prompt_handler(
            template: PromptTemplate,
        ) -> Callable[..., Awaitable[str]]:
            async def prompt_handler(**kwargs: Any) -> str:
                try:
                    return template.template.format(**kwargs)
                except KeyError as e:
                    return f"Error: Parámetro requerido no encontrado: {e}"

            # Asignar metadata
            prompt_handler.__name__ = template.name
            prompt_handler.__doc__ = template.description
            return prompt_handler

        # Registrar el prompt con el decorador
        handler = create_prompt_handler(prompt_template)
        mcp.prompt(name=prompt_template.name, description=prompt_template.description)(
            handler
        )


"""
Prompt registration for MCP Office Word Server.

This module provides functionality to register and manage prompt templates
for the MCP Office Word Server application.
"""
