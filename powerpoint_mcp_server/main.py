from mcp.server.fastmcp import FastMCP
import platform
import sys
from typing import Dict, List, Any
from .powerpoint_adapter import get_powerpoint_adapter

mcp = FastMCP("powerpoint-mcp")

# Get the appropriate PowerPoint adapter for the current platform
ppt_adapter = get_powerpoint_adapter()


@mcp.tool()
def initialize_powerpoint() -> Dict[str, Any]:
    """
    Initialize connection to PowerPoint.

    Returns:
        Status of initialization and platform information
    """
    success = ppt_adapter.initialize()
    return {
        "success": success,
        "platform": platform.system(),
        "adapter_type": type(ppt_adapter).__name__,
        "message": (
            "PowerPoint connection initialized"
            if success
            else "Failed to initialize PowerPoint"
        ),
    }


@mcp.tool()
def get_presentations() -> List[Dict[str, Any]]:
    """Get a list of all open PowerPoint presentations with their metadata."""
    return ppt_adapter.get_open_presentations()


@mcp.tool()
def open_presentation(path: str) -> Dict[str, Any]:
    """
    Open a PowerPoint presentation from the specified path.

    Args:
        path: Full path to the PowerPoint file (.pptx, .ppt)

    Returns:
        Dictionary with presentation ID and metadata
    """
    return ppt_adapter.open_presentation(path)


@mcp.tool()
def get_slides(presentation_id: str) -> List[Dict[str, Any]]:
    """
    Get a list of all slides in a presentation.

    Args:
        presentation_id: ID of the presentation

    Returns:
        List of slide metadata
    """
    return ppt_adapter.get_slides(presentation_id)


@mcp.tool()
def get_slide_text(presentation_id: str, slide_id: int) -> Dict[str, Any]:
    """
    Get all text content in a slide.

    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (integer)

    Returns:
        Dictionary containing text content organized by shape
    """
    return ppt_adapter.get_slide_text(presentation_id, slide_id)


@mcp.tool()
def update_text(
    presentation_id: str, slide_id: str, shape_id: str, text: str
) -> Dict[str, Any]:
    """
    Update the text content of a shape.

    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (numeric string)
        shape_id: ID of the shape (numeric string)
        text: New text content

    Returns:
        Status of the operation
    """
    return ppt_adapter.update_text(presentation_id, slide_id, shape_id, text)


@mcp.tool()
def save_presentation(presentation_id: str, path: str = None) -> Dict[str, Any]:
    """
    Save a presentation to disk.

    Args:
        presentation_id: ID of the presentation
        path: Optional path to save the file (if None, save to current location)

    Returns:
        Status of the operation
    """
    return ppt_adapter.save_presentation(presentation_id, path)


@mcp.tool()
def close_presentation(presentation_id: str, save: bool = True) -> Dict[str, Any]:
    """
    Close a presentation.

    Args:
        presentation_id: ID of the presentation
        save: Whether to save changes before closing

    Returns:
        Status of the operation
    """
    return ppt_adapter.close_presentation(presentation_id, save)


@mcp.tool()
def create_presentation() -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation.

    Returns:
        Dictionary containing new presentation ID and metadata
    """
    return ppt_adapter.create_presentation()


@mcp.tool()
def add_slide(presentation_id: str, layout_type: int = 1) -> Dict[str, Any]:
    """
    Add a new slide to the presentation.

    Args:
        presentation_id: ID of the presentation
        layout_type: Slide layout type (default is 1, title slide)
            1: ppLayoutTitle (title slide)
            2: ppLayoutText (slide with title and text)
            3: ppLayoutTwoColumns (two-column slide)
            7: ppLayoutBlank (blank slide)
            etc...

    Returns:
        Information about the new slide
    """
    return ppt_adapter.add_slide(presentation_id, layout_type)


@mcp.tool()
def add_text_box(
    presentation_id: str,
    slide_id: str,
    text: str,
    left: float = 100,
    top: float = 100,
    width: float = 400,
    height: float = 200,
) -> Dict[str, Any]:
    """
    Add a text box to a slide and set its text content.

    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (numeric string)
        text: Text content
        left: Left edge position of the text box (points)
        top: Top edge position of the text box (points)
        width: Width of the text box (points)
        height: Height of the text box (points)

    Returns:
        Operation status and ID of the new shape
    """
    return ppt_adapter.add_text_box(
        presentation_id, slide_id, text, left, top, width, height
    )


@mcp.tool()
def set_slide_title(presentation_id: str, slide_id: str, title: str) -> Dict[str, Any]:
    """
    Set the title text of a slide.

    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (numeric string)
        title: New title text

    Returns:
        Status of the operation
    """
    return ppt_adapter.set_slide_title(presentation_id, slide_id, title)


@mcp.tool()
def get_platform_info() -> Dict[str, Any]:
    """
    Get information about the current platform and available PowerPoint adapters.

    Returns:
        Platform and adapter information
    """
    return {
        "platform": platform.system(),
        "platform_release": platform.release(),
        "python_version": sys.version,
        "adapter_type": type(ppt_adapter).__name__,
        "adapter_available": hasattr(ppt_adapter, "available")
        and getattr(ppt_adapter, "available", False),
    }


def main():
    """Main entry point for the MCP server"""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
