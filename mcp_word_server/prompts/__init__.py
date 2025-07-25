"""
Prompts module for MCP Office Word Server.

This module contains prompt templates and registration functionality
for academic document generation.
"""

from .prompts import (
    PROMPT_REGISTRY,
    get_prompt_template,
    list_available_prompts,
    get_prompts_by_category,
    get_prompts_by_academic_level,
    PromptTemplate
)

from .register_prompts import register_prompts

__all__ = [
    'PROMPT_REGISTRY',
    'get_prompt_template',
    'list_available_prompts',
    'get_prompts_by_category',
    'get_prompts_by_academic_level',
    'PromptTemplate',
    'register_prompts'
]
