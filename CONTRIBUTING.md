# Contributing to Cross-Platform PowerPoint MCP Server

Thank you for your interest in contributing! This guide will help you get started.

## 🚀 Quick Start

1. **Fork the repository** on GitHub
2. **Clone your fork** locally:
   ```bash
   git clone https://github.com/YOUR_USERNAME/powerpoint-mcp-server.git
   cd powerpoint-mcp-server
   ```
3. **Create a virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
4. **Install in development mode**:
   ```bash
   pip install -e ".[dev]"
   ```

## 🔧 Development Workflow

### Creating a Feature Branch
```bash
git checkout -b feature/your-feature-name
```

### Running Tests
```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=powerpoint_mcp_server

# Run specific test file
pytest tests/test_adapter.py
```

### Code Style
We use:
- **Black** for code formatting
- **Flake8** for linting
- **Type hints** for better code documentation

```bash
# Format code
black powerpoint_mcp_server/

# Check linting
flake8 powerpoint_mcp_server/
```

## 📝 Contribution Types

### 🐛 Bug Reports
- Use GitHub Issues
- Include Python version, OS, and reproduction steps
- Provide error messages and stack traces

### 💡 Feature Requests
- Describe the use case and benefit
- Consider cross-platform compatibility
- Discuss in GitHub Discussions first for major features

### 🔧 Code Contributions
- Follow existing code patterns
- Add tests for new functionality
- Update documentation as needed
- Ensure cross-platform compatibility

## 🏗️ Architecture Guidelines

### Adapter Pattern
- New PowerPoint backends should implement `PowerPointAdapter`
- Platform detection in `get_powerpoint_adapter()`
- Graceful fallbacks for unsupported platforms

### Cross-Platform Considerations
- Test on multiple platforms when possible
- Use appropriate file paths (`pathlib` recommended)
- Handle platform-specific dependencies gracefully

### MCP Integration
- All tools should be properly decorated with `@mcp.tool()`
- Provide clear docstrings with parameter descriptions
- Return consistent response formats

## 🧪 Testing

### Test Categories
- **Unit tests**: Individual adapter functions
- **Integration tests**: Full MCP server workflow  
- **Platform tests**: Cross-platform compatibility

### Writing Tests
```python
import pytest
from powerpoint_mcp_server.powerpoint_adapter import get_powerpoint_adapter

def test_adapter_selection():
    adapter = get_powerpoint_adapter()
    assert adapter is not None
    assert hasattr(adapter, 'initialize')
```

## 📚 Documentation

### README Updates
- Keep installation instructions current
- Update feature lists for new capabilities
- Maintain platform compatibility matrix

### Code Documentation
- Use clear, descriptive docstrings
- Include parameter types and descriptions
- Document return value formats

### Examples
- Provide usage examples for new features
- Update existing examples if behavior changes

## 🚀 Pull Request Process

1. **Ensure tests pass** on your local machine
2. **Write descriptive commit messages**:
   ```
   feat: add support for slide animations
   
   - Implement animation API in both adapters
   - Add tests for animation functionality  
   - Update documentation with examples
   ```
3. **Create pull request** with:
   - Clear description of changes
   - Link to related issues
   - Screenshots/examples if UI-related
4. **Respond to feedback** and make requested changes
5. **Squash commits** if requested before merge

## 🏷️ Commit Message Format

Use conventional commits format:
- `feat:` - New features
- `fix:` - Bug fixes  
- `docs:` - Documentation changes
- `test:` - Adding/updating tests
- `refactor:` - Code refactoring
- `chore:` - Maintenance tasks

## 🌍 Platform Testing

### Windows
- Test both COM API and python-pptx paths
- Verify PowerPoint integration
- Test file path handling

### macOS/Linux  
- Ensure python-pptx functionality
- Test file operations
- Verify error handling

## 📋 Release Process

1. Update version in `pyproject.toml`
2. Update `CHANGELOG.md` 
3. Create release PR
4. Tag release after merge
5. GitHub Actions handles PyPI publishing

## ❓ Questions?

- **Discussions**: Use GitHub Discussions for questions
- **Issues**: Create issues for bugs or feature requests
- **Discord**: Join the Claude Desktop community

## 📜 Code of Conduct

By participating, you agree to:
- Be respectful and inclusive
- Focus on constructive feedback
- Help maintain a welcoming community
- Follow GitHub's Community Guidelines

Thank you for contributing! 🎉