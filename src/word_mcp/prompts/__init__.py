"""
Prompts module for MCP Office Word Server.

This module contains prompt templates and registration functionality
for academic document generation.
"""

from .prompts import (
    PROMPT_REGISTRY,
    PromptTemplate,
    get_prompt_template,
    get_prompts_by_academic_level,
    get_prompts_by_category,
    list_available_prompts,
)
from .register_prompts import register_prompts

__all__ = [
    "PROMPT_REGISTRY",
    "get_prompt_template",
    "list_available_prompts",
    "get_prompts_by_category",
    "get_prompts_by_academic_level",
    "PromptTemplate",
    "register_prompts",
]
