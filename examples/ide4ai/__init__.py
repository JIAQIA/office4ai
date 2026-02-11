"""
AI IDE - A powerful IDE environment designed for AI agents to interact with code.

This package provides a comprehensive IDE environment that AI agents can use to:
- Navigate and understand code structure
- Edit files with LSP support
- Execute commands in terminal environments
- Manage workspaces and projects

Main Components:
- IDE: Base IDE environment class
- PythonIDE: Python-specific IDE implementation
- BaseWorkspace: Workspace management
- BaseTerminalEnv: Terminal environment interface
"""

from ide4ai.base import IDE
from ide4ai.exceptions import IDEExecutionError, IDEProtocolError
from ide4ai.ides import IDESingleton, PyIDESingleton
from ide4ai.python_ide.ide import PythonIDE
from ide4ai.schema import IDEAction, IDEObs, LanguageId

__version__ = "0.1.0b0"

__all__ = [
    "IDE",
    "PythonIDE",
    "IDESingleton",
    "PyIDESingleton",
    "IDEAction",
    "IDEObs",
    "LanguageId",
    "IDEExecutionError",
    "IDEProtocolError",
]
