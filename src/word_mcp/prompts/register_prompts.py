from mcp.server.fastmcp import FastMCP

from word_mcp.prompts.prompts import PROMPT_REGISTRY


def register_prompts(mcp: FastMCP):
    """Register all prompts with the MCP server."""

    def create_prompt_function(template):
        async def prompt_function(**kwargs):
            return template.template.format(**kwargs)

        return prompt_function

    for prompt_name, prompt_template in PROMPT_REGISTRY.items():
        prompt_func = create_prompt_function(prompt_template)
        mcp.prompt(name=prompt_template.name)(prompt_func)


"""
Prompt registration for MCP Office Word Server.

This module provides functionality to register and manage prompt templates
for the MCP Office Word Server application.
"""
