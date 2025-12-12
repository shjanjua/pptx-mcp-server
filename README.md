# 📊 PPTX MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to create, edit, and manipulate PowerPoint presentations programmatically.

## ✨ Features

- **Full PowerPoint Control** - Create, read, and modify `.pptx` files without needing PowerPoint installed
- **AI-Native Design** - Built specifically for LLMs to generate and edit presentations through structured JSON
- **Rich Formatting Support** - Text styling, colors, alignment, bullets, shapes, and backgrounds
- **Template Workflows** - Extract content from existing presentations, modify, and regenerate
- **Visual Debugging** - Generate thumbnail grids to preview slides programmatically
- **Office XML Access** - Direct access to underlying OOXML for advanced customization

## 🚀 Quick Start

### 1. Install

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/pptx-mcp-server.git
cd pptx-mcp-server

# Install the package
pip install -e .
```

### 2. Configure Your MCP Client

Add to your MCP settings:

<details>
<summary><b>Claude Desktop</b></summary>

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

```json
{
  "mcpServers": {
    "pptx": {
      "command": "python",
      "args": ["-m", "pptx_mcp_server"]
    }
  }
}
```
</details>

<details>
<summary><b>Cursor</b></summary>

Add to your MCP configuration:

```json
{
  "mcpServers": {
    "pptx": {
      "command": "python",
      "args": ["-m", "pptx_mcp_server"]
    }
  }
}
```
</details>

<details>
<summary><b>Using a Virtual Environment (conda/venv)</b></summary>

Specify the full path to your Python interpreter:

```json
{
  "mcpServers": {
    "pptx": {
      "command": "/path/to/your/python",
      "args": ["-m", "pptx_mcp_server"]
    }
  }
}
```

Examples:
- Conda: `/Users/username/miniconda3/bin/python`
- venv: `/path/to/project/.venv/bin/python`
</details>

### 3. Restart Your MCP Client

Restart Claude Desktop, Cursor, or your MCP client to load the server.

## 📦 Requirements

- **Python 3.10+**
- Dependencies (installed automatically):
  - `mcp` - Model Context Protocol SDK
  - `python-pptx` - PowerPoint file manipulation
  - `Pillow` - Image processing
  - `lxml` - XML parsing
  - `defusedxml` - Secure XML parsing

### Optional (for thumbnails)

```bash
# macOS
brew install --cask libreoffice
brew install poppler

# Ubuntu/Debian
sudo apt-get install libreoffice poppler-utils
```

## 🛠️ Available Tools

| Tool | Description |
|------|-------------|
| `create_presentation` | Create new presentations from scratch |
| `extract_text_inventory` | Extract text content with positions and formatting |
| `apply_text_replacements` | Replace text using JSON specifications |
| `rearrange_slides` | Duplicate, delete, and reorder slides |
| `create_thumbnail_grid` | Generate visual thumbnail grids |
| `unpack_office_document` | Extract Office files to editable XML |
| `pack_office_document` | Rebuild Office files from XML |
| `validate_office_document` | Validate document structure |

## 📖 Usage Examples

### Create a New Presentation

```json
{
  "output_path": "/path/to/presentation.pptx",
  "layout": "16:9",
  "slides": [
    {
      "background": "#0f172a",
      "shapes": [
        {
          "type": "textbox",
          "left": 0.5,
          "top": 3,
          "width": 12,
          "height": 1.5,
          "text": "Welcome to My Presentation",
          "font_size": 54,
          "bold": true,
          "color": "#ffffff",
          "alignment": "center"
        },
        {
          "type": "textbox",
          "left": 0.5,
          "top": 5,
          "width": 12,
          "height": 1,
          "text": "Subtitle goes here",
          "font_size": 24,
          "color": "#94a3b8",
          "alignment": "center"
        }
      ]
    },
    {
      "shapes": [
        {
          "type": "textbox",
          "left": 0.5,
          "top": 0.5,
          "width": 12,
          "height": 1,
          "text": "Key Points",
          "font_size": 36,
          "bold": true
        },
        {
          "type": "textbox",
          "left": 0.5,
          "top": 1.8,
          "width": 12,
          "height": 5,
          "paragraphs": [
            {"text": "First important point", "font_size": 24, "bullet": true},
            {"text": "Second important point", "font_size": 24, "bullet": true},
            {"text": "Third important point", "font_size": 24, "bullet": true}
          ]
        }
      ]
    }
  ]
}
```

**Supported shape types:**
- `textbox` - Text content
- `rectangle` - Rectangle (can contain text)
- `rounded_rectangle` - Rounded corners
- `oval` - Circle/ellipse
- `image` - Image file (use `path` property)
- `line` - Line connector

**Supported layouts:** `16:9`, `4:3`, `widescreen`, `standard`

### Extract Text Inventory

Get all text content from an existing presentation:

```json
{
  "pptx_path": "/path/to/presentation.pptx"
}
```

Returns structured JSON:
```json
{
  "slide-0": {
    "shape-0": {
      "left": 0.5,
      "top": 1.0,
      "width": 12.0,
      "height": 1.5,
      "paragraphs": [
        {"text": "Title Text", "font_size": 44.0, "bold": true}
      ]
    }
  }
}
```

### Replace Text Content

Modify text in an existing presentation:

```json
{
  "pptx_path": "/path/to/template.pptx",
  "output_path": "/path/to/output.pptx",
  "replacements_json": {
    "slide-0": {
      "shape-0": [
        {"text": "New Title", "font_size": 44, "bold": true}
      ]
    }
  }
}
```

### Rearrange Slides

Reorder, duplicate, or remove slides:

```json
{
  "template_path": "/path/to/template.pptx",
  "output_path": "/path/to/output.pptx",
  "slide_sequence": "0,2,2,1,3"
}
```

- `0,2,2,1,3` → Keep slide 0, duplicate slide 2, then slides 1 and 3
- Omit an index to delete that slide

### Unpack/Pack for XML Editing

```json
// Unpack to directory
{
  "office_file": "/path/to/document.pptx",
  "output_dir": "/path/to/unpacked"
}

// Pack back to file
{
  "input_dir": "/path/to/unpacked",
  "output_file": "/path/to/output.pptx"
}
```

## 🔧 Troubleshooting

<details>
<summary><b>ModuleNotFoundError: No module named 'mcp'</b></summary>

Ensure you're using the correct Python environment:

```bash
# Check which Python pip uses
pip --version

# Install with specific Python
/path/to/python -m pip install -e .
```
</details>

<details>
<summary><b>Server not appearing in MCP client</b></summary>

1. Verify the config file path is correct for your OS
2. Ensure JSON syntax is valid (no trailing commas)
3. Restart the MCP client completely
4. Check logs for errors
</details>

<details>
<summary><b>Thumbnail generation fails</b></summary>

Install LibreOffice and poppler:

```bash
# macOS
brew install --cask libreoffice && brew install poppler

# Linux
sudo apt-get install libreoffice poppler-utils
```
</details>

<details>
<summary><b>Permission denied errors</b></summary>

Ensure the output paths are writable and parent directories exist.
</details>

## 📁 Project Structure

```
pptx-mcp-server/
├── pyproject.toml              # Package configuration
├── README.md
└── pptx_mcp_server/
    ├── __init__.py
    ├── server.py               # MCP server implementation
    └── tools/
        ├── __init__.py
        ├── create.py           # Create new presentations
        ├── inventory.py        # Extract text content
        ├── replace.py          # Text replacement
        ├── rearrange.py        # Slide manipulation
        ├── thumbnail.py        # Visual thumbnails
        └── ooxml.py            # XML pack/unpack/validate
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/)
- Uses the [Model Context Protocol](https://modelcontextprotocol.io/) by Anthropic
