"""Security utilities for file path validation.

This module provides functions for validating file paths against
allowed directories for security purposes.
"""

import os


def get_allowed_directories() -> list[str]:
    """Get the list of allowed directories from environment variables.

    Returns:
        List of absolute paths to directories where documents can be created/accessed.
        Defaults to ['./documents'] if MCP_ALLOWED_DIRECTORIES is not set.
    """
    allowed_dirs_str = os.environ.get("MCP_ALLOWED_DIRECTORIES", "./documents")
    allowed_dirs = [dir.strip() for dir in allowed_dirs_str.split(",")]
    return [os.path.abspath(dir) for dir in allowed_dirs]


def is_path_in_allowed_directories(file_path: str) -> tuple[bool, str | None]:
    """Check if the given file path is within allowed directories.

    Args:
        file_path: The file path to validate.

    Returns:
        Tuple of (is_allowed, error_message) where is_allowed is a boolean
        indicating if the path is allowed, and error_message provides details
        if the path is not allowed.
    """
    allowed_dirs = get_allowed_directories()
    abs_path = os.path.abspath(file_path)

    for allowed_dir in allowed_dirs:
        try:
            if os.path.commonpath([allowed_dir, abs_path]) == allowed_dir:
                return True, None
        except ValueError:
            continue

    error_msg = (
        f"Path '{file_path}' is not in allowed directories: {', '.join(allowed_dirs)}"
    )
    return False, error_msg
