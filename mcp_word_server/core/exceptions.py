"""Custom exceptions for the MCP Word Server application.

This module defines custom exceptions used throughout the application
for better error handling and more informative error messages.
"""

class DocumentProcessingError(Exception):
    """Raised when there is an error processing a Word document.
    
    This exception should be used for errors that occur during document
    manipulation, such as adding content, formatting, or saving documents.
    """
    pass


class ValidationError(ValueError):
    """Raised when input validation fails.
    
    This exception should be used for errors related to invalid input
    parameters or data that doesn't meet the required criteria.
    """
    pass


class FileOperationError(IOError):
    """Raised when there is an error performing file operations.
    
    This includes errors related to reading, writing, or accessing files.
    """
    pass


class StyleError(Exception):
    """Raised when there is an error related to document styles.
    
    This includes missing styles, invalid style names, or style application errors.
    """
    pass


class ConfigurationError(Exception):
    """Raised when there is a configuration error in the application.
    
    This includes missing or invalid configuration settings.
    """
    pass
