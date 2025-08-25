# 🎯 Cross-Platform PowerPoint MCP Server

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PyPI version](https://badge.fury.io/py/powerpoint-mcp-server.svg)](https://badge.fury.io/py/powerpoint-mcp-server)
[![Cross-Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](https://github.com/your-username/powerpoint-mcp-server)

A **cross-platform** PowerPoint automation server that works with [Claude Desktop](https://claude.ai/) via the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/). Create, edit, and manage PowerPoint presentations on **Windows, macOS, and Linux** using AI assistance.

## ✨ Features

- 🌍 **Cross-platform compatibility** (Windows, macOS, Linux)
- 🎨 **Complete PowerPoint automation** - Create, edit, save presentations
- 🔄 **Intelligent adapter selection** - COM API on Windows, python-pptx everywhere else
- 🤖 **Claude Desktop integration** - Control PowerPoint through natural language
- 📝 **Rich text manipulation** - Add text boxes, update content, set titles
- 🎭 **Advanced techniques support** - Perfect for morph transitions and animations
- 📦 **Easy installation** - Available on PyPI

## 🚀 Quick Start

### Installation

```bash
pip install powerpoint-mcp-server
```

**Platform-specific notes:**
- **Windows**: Optionally install `pywin32` for direct PowerPoint COM API access
- **macOS/Linux**: Uses `python-pptx` library (installed automatically)

### Claude Desktop Configuration

Add to your Claude Desktop configuration file:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`  
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
**Linux**: `~/.config/claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["-m", "powerpoint_mcp_server"]
    }
  }
}
```

## 💬 Usage Examples

Once configured, interact with PowerPoint through Claude Desktop:

```
🤖 What platform am I running on and what PowerPoint adapter is available?

🤖 Please create a new PowerPoint presentation with a title slide called "AI-Powered Presentations"

🤖 Add a content slide explaining the benefits of cross-platform automation

🤖 Save the presentation to ~/Documents/my-ai-presentation.pptx
```

## 🔧 Available Functions

### Core Operations
- `initialize_powerpoint()` - Initialize PowerPoint connection
- `get_platform_info()` - Get system and adapter information
- `create_presentation()` - Create new presentation
- `open_presentation(path)` - Open existing presentation
- `save_presentation(id, path)` - Save presentation
- `close_presentation(id)` - Close presentation

### Slide Management
- `get_slides(presentation_id)` - List all slides
- `add_slide(presentation_id, layout_type)` - Add new slide
- `get_slide_text(presentation_id, slide_id)` - Extract slide text
- `set_slide_title(presentation_id, slide_id, title)` - Set slide title

### Content Editing
- `add_text_box(presentation_id, slide_id, text, ...)` - Add text box
- `update_text(presentation_id, slide_id, shape_id, text)` - Update text content

## 🖥️ Platform Support

| Feature | Windows + PowerPoint | Windows (python-pptx) | macOS | Linux |
|---------|---------------------|----------------------|--------|--------|
| Create presentations | ✅ | ✅ | ✅ | ✅ |
| Edit presentations | ✅ | ✅ | ✅ | ✅ |
| Live PowerPoint control | ✅ | ❌ | ❌ | ❌ |
| File-based operations | ✅ | ✅ | ✅ | ✅ |
| Morph transitions* | ✅ | ✅ | ✅ | ✅ |

*\*Morph transitions require PowerPoint Desktop for playback*

## 🎭 Advanced Use Cases

This MCP server is perfect for:
- **AI-assisted presentation creation**
- **Batch processing PowerPoint files**
- **Cross-platform presentation workflows**
- **Advanced animation techniques** (liquid masks, morph effects)
- **Automated content generation**
- **Educational presentation tools**

## 🏗️ Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│  Claude Desktop │◄──►│  MCP Protocol    │◄──►│   This Server   │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                                         │
                                                         ▼
                                               ┌─────────────────┐
                                               │ Platform Detect │
                                               └─────────────────┘
                                                         │
                                    ┌────────────────────┼────────────────────┐
                                    ▼                    ▼                    ▼
                          ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
                          │ Windows COM API │  │   python-pptx   │  │ Fallback Handler│
                          │   (pywin32)     │  │ (Cross-platform)│  │   (No adapter)  │
                          └─────────────────┘  └─────────────────┘  └─────────────────┘
```

## 🛠️ Development

### Setup Development Environment

```bash
# Clone repository
git clone https://github.com/your-username/powerpoint-mcp-server.git
cd powerpoint-mcp-server

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in development mode
pip install -e ".[dev]"
```

### Running Tests

```bash
pytest
```

### Building Package

```bash
python -m build
```

## 📝 Requirements

- **Python 3.10+**
- **Claude Desktop** client
- **Optional**: PowerPoint Desktop (for live control on Windows)

## 🤝 Contributing

Contributions are welcome! Please see our [Contributing Guide](CONTRIBUTING.md).

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes and add tests
4. Commit: `git commit -m 'Add amazing feature'`
5. Push: `git push origin feature/amazing-feature`
6. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) - Communication protocol
- [python-pptx](https://python-pptx.readthedocs.io/) - Cross-platform PowerPoint library  
- [pywin32](https://github.com/mhammond/pywin32) - Windows COM API access
- [Claude Desktop](https://claude.ai/) - AI-powered automation platform

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/your-username/powerpoint-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/your-username/powerpoint-mcp-server/discussions)
- **Documentation**: [Project Wiki](https://github.com/your-username/powerpoint-mcp-server/wiki)

---

**Made with ❤️ for the Claude Desktop community**