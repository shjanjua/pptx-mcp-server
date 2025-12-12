"""
PPTX MCP Server - Main server implementation.
"""

import asyncio
import json
import logging
from pathlib import Path
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from .tools.create import create_presentation
from .tools.inventory import extract_text_inventory, get_inventory_as_dict
from .tools.replace import apply_replacements
from .tools.rearrange import rearrange_presentation
from .tools.thumbnail import create_thumbnail_grids
from .tools.ooxml import unpack_document, pack_document, validate_document

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

server = Server("pptx-mcp-server")


@server.list_tools()
async def list_tools() -> list[Tool]:
    """List all available tools."""
    return [
        Tool(
            name="create_presentation",
            description=(
                "Create a new PowerPoint presentation from scratch. "
                "Accepts a JSON specification with slides, shapes, text content, and formatting. "
                "Supports text boxes, rectangles, ovals, images, and various text formatting options. "
                "Use this to create new presentations before using other tools to modify them."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "output_path": {
                        "type": "string",
                        "description": "Path to save the new PowerPoint file (.pptx)",
                    },
                    "layout": {
                        "type": "string",
                        "description": "Slide layout: '16:9' (default), '4:3', 'widescreen', or 'standard'",
                        "default": "16:9",
                    },
                    "slides": {
                        "type": "array",
                        "description": "Array of slide specifications",
                        "items": {
                            "type": "object",
                            "properties": {
                                "background": {
                                    "type": "string",
                                    "description": "Background color (hex, e.g., '#FFFFFF')",
                                },
                                "shapes": {
                                    "type": "array",
                                    "description": "Array of shape specifications",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "type": {
                                                "type": "string",
                                                "enum": ["textbox", "rectangle", "rounded_rectangle", "oval", "image", "line"],
                                                "description": "Shape type",
                                            },
                                            "left": {"type": "number", "description": "Left position in inches"},
                                            "top": {"type": "number", "description": "Top position in inches"},
                                            "width": {"type": "number", "description": "Width in inches"},
                                            "height": {"type": "number", "description": "Height in inches"},
                                            "text": {"type": "string", "description": "Text content (for simple text)"},
                                            "paragraphs": {
                                                "type": "array",
                                                "description": "Array of paragraph specs for multi-paragraph text",
                                                "items": {
                                                    "type": "object",
                                                    "properties": {
                                                        "text": {"type": "string"},
                                                        "font_size": {"type": "number"},
                                                        "font_name": {"type": "string"},
                                                        "bold": {"type": "boolean"},
                                                        "italic": {"type": "boolean"},
                                                        "color": {"type": "string"},
                                                        "alignment": {"type": "string", "enum": ["left", "center", "right"]},
                                                        "bullet": {"type": "boolean"},
                                                    },
                                                },
                                            },
                                            "font_size": {"type": "number", "description": "Font size in points"},
                                            "font_name": {"type": "string", "description": "Font name"},
                                            "bold": {"type": "boolean"},
                                            "italic": {"type": "boolean"},
                                            "color": {"type": "string", "description": "Text color (hex)"},
                                            "fill": {"type": "string", "description": "Shape fill color (hex)"},
                                            "alignment": {"type": "string", "enum": ["left", "center", "right"]},
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
                "required": ["output_path"],
            },
        ),
        Tool(
            name="extract_text_inventory",
            description=(
                "Extract structured text content from a PowerPoint presentation. "
                "Returns JSON with all text shapes, their positions, and formatting details. "
                "Useful for understanding presentation structure before making replacements."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Path to the PowerPoint file (.pptx)",
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Optional: Path to save the inventory JSON file",
                    },
                    "issues_only": {
                        "type": "boolean",
                        "description": "If true, only include shapes with overflow or overlap issues",
                        "default": False,
                    },
                },
                "required": ["pptx_path"],
            },
        ),
        Tool(
            name="apply_text_replacements",
            description=(
                "Apply text replacements to a PowerPoint presentation using a JSON specification. "
                "The JSON should map slide/shape IDs to new paragraph content with formatting. "
                "All text shapes are cleared unless explicitly provided with new content."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Path to the input PowerPoint file",
                    },
                    "replacements_json": {
                        "type": "string",
                        "description": "Path to JSON file with replacement specifications, or inline JSON string",
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Path for the output PowerPoint file",
                    },
                },
                "required": ["pptx_path", "replacements_json", "output_path"],
            },
        ),
        Tool(
            name="rearrange_slides",
            description=(
                "Rearrange slides in a PowerPoint presentation. "
                "Can duplicate, delete, and reorder slides based on a sequence of indices. "
                "Slide indices are 0-based. The same index can appear multiple times to duplicate."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "template_path": {
                        "type": "string",
                        "description": "Path to the template/input PowerPoint file",
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Path for the output PowerPoint file",
                    },
                    "slide_sequence": {
                        "type": "string",
                        "description": "Comma-separated slide indices (0-based), e.g., '0,34,34,50,52'",
                    },
                },
                "required": ["template_path", "output_path", "slide_sequence"],
            },
        ),
        Tool(
            name="create_thumbnail_grid",
            description=(
                "Create visual thumbnail grids from PowerPoint slides. "
                "Useful for quick visual analysis of presentation structure and layouts. "
                "For large presentations, multiple grid images are created automatically."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Path to the PowerPoint file",
                    },
                    "output_prefix": {
                        "type": "string",
                        "description": "Output prefix for image files (default: 'thumbnails')",
                        "default": "thumbnails",
                    },
                    "cols": {
                        "type": "integer",
                        "description": "Number of columns in the grid (3-6, default: 5)",
                        "default": 5,
                        "minimum": 3,
                        "maximum": 6,
                    },
                    "outline_placeholders": {
                        "type": "boolean",
                        "description": "Outline text placeholders with red borders",
                        "default": False,
                    },
                },
                "required": ["pptx_path"],
            },
        ),
        Tool(
            name="unpack_office_document",
            description=(
                "Unpack an Office document (.docx, .pptx, .xlsx) to a directory. "
                "The XML files are pretty-printed for easy reading and editing. "
                "Use this to inspect or manually edit the raw XML structure."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "office_file": {
                        "type": "string",
                        "description": "Path to the Office file (.docx, .pptx, or .xlsx)",
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "Directory to extract contents to",
                    },
                },
                "required": ["office_file", "output_dir"],
            },
        ),
        Tool(
            name="pack_office_document",
            description=(
                "Pack a directory back into an Office document (.docx, .pptx, .xlsx). "
                "Removes pretty-printing whitespace from XML before packing. "
                "Can optionally validate the document before saving."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "input_dir": {
                        "type": "string",
                        "description": "Directory containing unpacked Office document",
                    },
                    "output_file": {
                        "type": "string",
                        "description": "Path for the output Office file",
                    },
                    "validate": {
                        "type": "boolean",
                        "description": "Validate the document after packing (requires LibreOffice)",
                        "default": False,
                    },
                    "force": {
                        "type": "boolean",
                        "description": "Skip validation and pack anyway",
                        "default": False,
                    },
                },
                "required": ["input_dir", "output_file"],
            },
        ),
        Tool(
            name="validate_office_document",
            description=(
                "Validate an unpacked Office document against XSD schemas. "
                "Checks XML well-formedness, namespace declarations, unique IDs, "
                "file references, content types, and schema compliance. "
                "Returns detailed error messages for any issues found."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "unpacked_dir": {
                        "type": "string",
                        "description": "Path to unpacked Office document directory",
                    },
                    "original_file": {
                        "type": "string",
                        "description": "Path to original Office file for comparison",
                    },
                    "verbose": {
                        "type": "boolean",
                        "description": "Enable verbose output",
                        "default": False,
                    },
                },
                "required": ["unpacked_dir", "original_file"],
            },
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls."""
    try:
        if name == "create_presentation":
            result = await handle_create_presentation(arguments)
        elif name == "extract_text_inventory":
            result = await handle_extract_inventory(arguments)
        elif name == "apply_text_replacements":
            result = await handle_apply_replacements(arguments)
        elif name == "rearrange_slides":
            result = await handle_rearrange_slides(arguments)
        elif name == "create_thumbnail_grid":
            result = await handle_create_thumbnails(arguments)
        elif name == "unpack_office_document":
            result = await handle_unpack_document(arguments)
        elif name == "pack_office_document":
            result = await handle_pack_document(arguments)
        elif name == "validate_office_document":
            result = await handle_validate_document(arguments)
        else:
            result = f"Unknown tool: {name}"

        return [TextContent(type="text", text=result)]

    except Exception as e:
        logger.exception(f"Error in tool {name}")
        return [TextContent(type="text", text=f"Error: {str(e)}")]


async def handle_create_presentation(args: dict[str, Any]) -> str:
    """Handle create_presentation tool."""
    output_path = args.get("output_path")
    if not output_path:
        return "Error: output_path is required"

    layout = args.get("layout", "16:9")
    slides = args.get("slides", [])

    # If no slides provided, create a single blank slide
    if not slides:
        slides = [{"shapes": []}]

    try:
        result = create_presentation(
            output_path=output_path,
            layout=layout,
            slides=slides,
        )
        return result
    except Exception as e:
        return f"Error creating presentation: {str(e)}"


async def handle_extract_inventory(args: dict[str, Any]) -> str:
    """Handle extract_text_inventory tool."""
    pptx_path = Path(args["pptx_path"])
    output_path = args.get("output_path")
    issues_only = args.get("issues_only", False)

    if not pptx_path.exists():
        return f"Error: File not found: {pptx_path}"

    if not pptx_path.suffix.lower() == ".pptx":
        return "Error: Input must be a PowerPoint file (.pptx)"

    inventory = get_inventory_as_dict(pptx_path, issues_only=issues_only)

    if output_path:
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        with open(output, "w", encoding="utf-8") as f:
            json.dump(inventory, f, indent=2, ensure_ascii=False)

        total_slides = len(inventory)
        total_shapes = sum(len(shapes) for shapes in inventory.values())
        return f"Inventory saved to: {output_path}\nFound text in {total_slides} slides with {total_shapes} text elements"
    else:
        return json.dumps(inventory, indent=2, ensure_ascii=False)


async def handle_apply_replacements(args: dict[str, Any]) -> str:
    """Handle apply_text_replacements tool."""
    pptx_path = Path(args["pptx_path"])
    replacements_input = args["replacements_json"]
    output_path = args["output_path"]

    if not pptx_path.exists():
        return f"Error: Input file not found: {pptx_path}"

    # Check if replacements_input is a file path or inline JSON
    replacements_path = Path(replacements_input)
    if replacements_path.exists():
        json_path = str(replacements_path)
    else:
        # Assume it's inline JSON - write to temp file
        import tempfile
        try:
            replacements_data = json.loads(replacements_input)
            with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
                json.dump(replacements_data, f)
                json_path = f.name
        except json.JSONDecodeError:
            return f"Error: Invalid JSON in replacements_json"

    try:
        apply_replacements(str(pptx_path), json_path, output_path)
        return f"Successfully applied replacements. Output saved to: {output_path}"
    except ValueError as e:
        return f"Validation error: {str(e)}"
    except Exception as e:
        return f"Error applying replacements: {str(e)}"


async def handle_rearrange_slides(args: dict[str, Any]) -> str:
    """Handle rearrange_slides tool."""
    template_path = Path(args["template_path"])
    output_path = Path(args["output_path"])
    sequence_str = args["slide_sequence"]

    if not template_path.exists():
        return f"Error: Template file not found: {template_path}"

    try:
        slide_sequence = [int(x.strip()) for x in sequence_str.split(",")]
    except ValueError:
        return "Error: Invalid sequence format. Use comma-separated integers (e.g., 0,34,34,50,52)"

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        rearrange_presentation(template_path, output_path, slide_sequence)
        return f"Successfully rearranged slides. Output saved to: {output_path}"
    except ValueError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error rearranging slides: {str(e)}"


async def handle_create_thumbnails(args: dict[str, Any]) -> str:
    """Handle create_thumbnail_grid tool."""
    pptx_path = Path(args["pptx_path"])
    output_prefix = args.get("output_prefix", "thumbnails")
    cols = args.get("cols", 5)
    outline_placeholders = args.get("outline_placeholders", False)

    if not pptx_path.exists():
        return f"Error: File not found: {pptx_path}"

    if not pptx_path.suffix.lower() == ".pptx":
        return "Error: Input must be a PowerPoint file (.pptx)"

    cols = min(max(cols, 3), 6)  # Clamp to 3-6

    try:
        grid_files = create_thumbnail_grids(
            pptx_path,
            output_prefix,
            cols=cols,
            outline_placeholders=outline_placeholders
        )
        return f"Created {len(grid_files)} grid(s):\n" + "\n".join(f"  - {f}" for f in grid_files)
    except Exception as e:
        return f"Error creating thumbnails: {str(e)}"


async def handle_unpack_document(args: dict[str, Any]) -> str:
    """Handle unpack_office_document tool."""
    office_file = Path(args["office_file"])
    output_dir = Path(args["output_dir"])

    if not office_file.exists():
        return f"Error: File not found: {office_file}"

    try:
        result = unpack_document(office_file, output_dir)
        return result
    except Exception as e:
        return f"Error unpacking document: {str(e)}"


async def handle_pack_document(args: dict[str, Any]) -> str:
    """Handle pack_office_document tool."""
    input_dir = Path(args["input_dir"])
    output_file = Path(args["output_file"])
    validate = args.get("validate", False)
    force = args.get("force", False)

    if not input_dir.is_dir():
        return f"Error: Directory not found: {input_dir}"

    try:
        success = pack_document(input_dir, output_file, validate=validate and not force)
        if success:
            msg = f"Successfully packed document: {output_file}"
            if force:
                msg += "\nWarning: Validation was skipped"
            return msg
        else:
            return "Error: Validation failed. Document may be corrupt. Use force=true to pack anyway."
    except ValueError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error packing document: {str(e)}"


async def handle_validate_document(args: dict[str, Any]) -> str:
    """Handle validate_office_document tool."""
    unpacked_dir = Path(args["unpacked_dir"])
    original_file = Path(args["original_file"])
    verbose = args.get("verbose", False)

    if not unpacked_dir.is_dir():
        return f"Error: Directory not found: {unpacked_dir}"

    if not original_file.exists():
        return f"Error: Original file not found: {original_file}"

    try:
        success, messages = validate_document(unpacked_dir, original_file, verbose=verbose)
        result = "\n".join(messages)
        if success:
            return f"All validations PASSED!\n{result}" if result else "All validations PASSED!"
        else:
            return f"Validation FAILED:\n{result}"
    except Exception as e:
        return f"Error validating document: {str(e)}"


def main():
    """Main entry point."""
    asyncio.run(run_server())


async def run_server():
    """Run the MCP server."""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    main()
