"""Tests for PowerPoint adapter functionality."""

import pytest
import platform
from powerpoint_mcp_server.powerpoint_adapter import (
    get_powerpoint_adapter,
    CrossPlatformPPTXAdapter,
    WindowsCOMAdapter,
)


def test_adapter_selection():
    """Test that the adapter selection returns a valid adapter."""
    adapter = get_powerpoint_adapter()
    assert adapter is not None
    assert hasattr(adapter, "initialize")
    assert hasattr(adapter, "create_presentation")


def test_adapter_initialization():
    """Test adapter initialization."""
    adapter = get_powerpoint_adapter()
    # Should not raise an exception
    result = adapter.initialize()
    assert isinstance(result, bool)


def test_cross_platform_adapter_available():
    """Test that CrossPlatformPPTXAdapter is available."""
    adapter = CrossPlatformPPTXAdapter()
    # Should have python-pptx available in CI environment
    assert hasattr(adapter, "available")


def test_windows_adapter_behavior():
    """Test Windows COM adapter behavior on different platforms."""
    adapter = WindowsCOMAdapter()
    if platform.system().lower() == "windows":
        # On Windows, adapter might be available
        assert hasattr(adapter, "available")
    else:
        # On non-Windows, COM adapter should not be available
        assert hasattr(adapter, "available")
        assert not adapter.available


def test_adapter_interface():
    """Test that adapters implement the required interface."""
    adapter = get_powerpoint_adapter()
    
    # Test required methods exist
    required_methods = [
        "initialize",
        "get_open_presentations", 
        "open_presentation",
        "create_presentation",
        "save_presentation",
        "close_presentation",
        "get_slides",
        "add_slide",
        "get_slide_text",
        "update_text",
        "add_text_box",
        "set_slide_title",
    ]
    
    for method in required_methods:
        assert hasattr(adapter, method), f"Adapter missing method: {method}"
        assert callable(getattr(adapter, method)), f"Method {method} is not callable"


def test_platform_detection():
    """Test platform detection works correctly."""
    current_platform = platform.system().lower()
    adapter = get_powerpoint_adapter()
    
    # Verify we got an appropriate adapter for the platform
    if current_platform == "windows":
        # Windows should get either COM or python-pptx adapter
        assert isinstance(adapter, (WindowsCOMAdapter, CrossPlatformPPTXAdapter))
    else:
        # macOS/Linux should get python-pptx adapter
        assert isinstance(adapter, CrossPlatformPPTXAdapter)


@pytest.mark.integration  
def test_presentation_creation():
    """Integration test for presentation creation."""
    adapter = get_powerpoint_adapter()
    
    if not adapter.initialize():
        pytest.skip("Adapter initialization failed")
    
    # Test creating a presentation
    result = adapter.create_presentation()
    
    # Should either succeed or fail gracefully
    assert isinstance(result, dict)
    if "error" not in result:
        assert "id" in result
        assert "name" in result
        assert "slide_count" in result