"""File utility functions for the Word Document Server.

This module provides helper functions for file system operations such as
checking permissions, copying files, and ensuring correct file extensions.
"""

import asyncio
import functools
import inspect
import os
from pathlib import Path
from typing import Any, Callable, Dict, Optional, Tuple, TypeVar, Union, cast

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

F = TypeVar("F", bound=Callable[..., Any])


# Decorator function to validate a .docx file path
def validate_docx_file(param_name: str) -> Callable[[F], F]:
    """A universal decorator to validate a .docx file path.

    This decorator ensures that the function argument specified by `param_name`
    is a valid, existing .docx file. It works for both sync and async functions.

    Validation Steps:
        - Argument exists in the function call.
        - The path points to an existing file.
        - The file has a .docx extension.
        - The file is a valid (non-corrupt) Word document.
    """

    def decorator(func: F) -> F:
        if asyncio.iscoroutinefunction(func):

            @functools.wraps(func)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                path_value = _get_argument_value(func, param_name, args, kwargs)
                if path_value is None:
                    return {"error": f"Parameter '{param_name}' not found."}

                if error := _validate_docx_path(path_value):
                    return error

                return await func(*args, **kwargs)

            return cast(F, async_wrapper)
        else:

            @functools.wraps(func)
            def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
                path_value = _get_argument_value(func, param_name, args, kwargs)
                if path_value is None:
                    return {"error": f"Parameter '{param_name}' not found."}

                if error := _validate_docx_path(path_value):
                    return error

                return func(*args, **kwargs)

            return cast(F, sync_wrapper)

    return decorator


def check_file_writeable(param_name: str) -> Callable[[F], F]:
    """Decorador que verifica si el archivo indicado es escribible."""

    def decorator(func: F) -> F:
        if inspect.iscoroutinefunction(func):

            @functools.wraps(func)
            async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
                value = _get_argument_value(func, param_name, args, kwargs)
                if value is None:
                    return {"error": f"Missing parameter '{param_name}'."}

                ok, error_msg = _check_file_writeable(value)
                if not ok:
                    return {"error": f"Cannot write to file: {error_msg}"}

                return await func(*args, **kwargs)

            return cast(F, async_wrapper)

        else:

            @functools.wraps(func)
            def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
                value = _get_argument_value(func, param_name, args, kwargs)
                if value is None:
                    return {"error": f"Missing parameter '{param_name}'."}

                ok, error_msg = _check_file_writeable(value)
                if not ok:
                    return {"error": f"Cannot write to file: {error_msg}"}

                return func(*args, **kwargs)

            return cast(F, sync_wrapper)

    return decorator


# helper function to check if a file is writeable
def _check_file_writeable(path_value: Union[str, Path]) -> Tuple[bool, str]:
    """Checks if a file is writeable."""
    try:
        path = Path(path_value).resolve()
        if not path.exists():
            return False, f"File '{path}' does not exist."
        if not os.access(path, os.W_OK):
            return False, f"File '{path}' is not writeable."
        return True, ""
    except Exception as e:
        return False, str(e)


# helper function to extract an argument's value from args or kwargs
def _get_argument_value(
    func: Callable, name: str, args: tuple, kwargs: dict
) -> Optional[Any]:
    """Extracts an argument's value from args or kwargs."""
    if name in kwargs:
        return kwargs[name]

    try:
        sig = inspect.signature(func)
        param_names = list(sig.parameters.keys())
        if name in param_names:
            index = param_names.index(name)
            if index < len(args):
                return args[index]
    except (ValueError, IndexError):
        return None
    return None


# helper function to validate a .docx file path of validate_docx_file decorator
def _validate_docx_path(path_str: str) -> Optional[Dict[str, str]]:
    """Performs validation checks on a given .docx file path."""

    path = Path(path_str).resolve()

    if not path.exists():
        return {"error": f"File '{path}' does not exist."}

    if path.suffix.lower() != ".docx":
        return {"error": f"File '{path}' is not a .docx document."}

    try:
        # Check if the file is a valid Word document by trying to open it.
        Document(str(path))
    except PackageNotFoundError:
        return {"error": f"File '{path}' is not a valid Word document (.docx)."}
    except Exception as e:
        return {"error": f"Could not open document: {e}"}

    return None
