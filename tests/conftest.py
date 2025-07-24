"""Configuration file for pytest.

This file contains fixtures and configuration that will be automatically loaded by pytest.
It's the recommended place to put fixtures that will be used across multiple test files.
"""
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import pytest
from pathlib import Path
from typing import Generator

# Add the project root to the Python path
PROJECT_ROOT = Path(__file__).parent

# This ensures that pytest can find your modules
import sys
sys.path.insert(0, str(PROJECT_ROOT))

# Fixtures can be defined here and will be automatically discovered by pytest
# For example, you could add fixtures for test data, database connections, etc.

# Example of a fixture that could be used across tests
@pytest.fixture(scope="session")
def test_data_dir() -> Path:
    """Return the path to the test data directory."""
    return PROJECT_ROOT / "test_data"
