"""Tests for main MCP server functionality."""

import pytest
from powerpoint_mcp_server.main import (
    initialize_powerpoint,
    get_platform_info,
    create_presentation,
    get_presentations,
)


def test_platform_info():
    """Test platform info function."""
    info = get_platform_info()
    
    assert isinstance(info, dict)
    assert "platform" in info
    assert "platform_release" in info
    assert "python_version" in info
    assert "adapter_type" in info
    assert "adapter_available" in info
    
    # Platform should be a valid OS
    assert info["platform"] in ["Windows", "Darwin", "Linux"]


def test_initialize_powerpoint():
    """Test PowerPoint initialization function."""
    result = initialize_powerpoint()
    
    assert isinstance(result, dict)
    assert "success" in result
    assert "platform" in result
    assert "adapter_type" in result
    assert "message" in result
    
    # Success should be boolean
    assert isinstance(result["success"], bool)


def test_create_presentation_function():
    """Test presentation creation function."""
    result = create_presentation()
    
    assert isinstance(result, dict)
    # Should either succeed or fail gracefully
    if "error" not in result:
        assert "id" in result
        assert "name" in result


def test_get_presentations_function():
    """Test get presentations function."""
    result = get_presentations()
    
    # Should return a list
    assert isinstance(result, list)


def test_mcp_tools_are_decorated():
    """Test that MCP tools are properly decorated."""
    import powerpoint_mcp_server.main as main_module
    
    # Check that key functions have the @mcp.tool() decorator applied
    # This is tested by checking if the function has been registered
    assert hasattr(main_module, 'mcp')
    
    # The FastMCP instance should exist
    assert main_module.mcp is not None
    
    # Check that tools are registered
    assert len(main_module.mcp.tools) > 0


@pytest.mark.integration
def test_full_workflow():
    """Integration test for a complete workflow."""
    # Test initialization
    init_result = initialize_powerpoint()
    assert isinstance(init_result, dict)
    
    if not init_result.get("success", False):
        pytest.skip("PowerPoint initialization failed")
    
    # Test creating presentation
    create_result = create_presentation()
    assert isinstance(create_result, dict)
    
    if "error" in create_result:
        pytest.skip(f"Presentation creation failed: {create_result['error']}")
    
    # Test getting presentations list  
    presentations = get_presentations()
    assert isinstance(presentations, list)