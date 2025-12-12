"""
PPTX MCP Server - PowerPoint presentation creation, editing, and analysis.

This MCP server provides tools for:
- Extracting text inventory from presentations
- Applying text replacements
- Rearranging slides (duplicate, delete, reorder)
- Creating thumbnail grids
- Unpacking/packing Office documents (OOXML)
- Validating Office document XML
"""

from .server import main

__all__ = ["main"]
