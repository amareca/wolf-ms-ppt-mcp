"""FastMCP server definition for wolf-ms-ppt-mcp (Phase 1)."""

from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from . import __version__
from .ops import images as image_ops
from .ops import layouts as layout_ops
from .ops import slides as slide_ops
from .ops import text as text_ops
from .state import PresentationState
from .units import emu_to_inches

mcp = MCPServer("wolf-ms-ppt-mcp", version=__version__)

_state = PresentationState()


def _deck_summary(prs) -> dict:
    return {
        "slide_count": len(prs.slides),
        "slide_width_in": emu_to_inches(prs.slide_width),
        "slide_height_in": emu_to_inches(prs.slide_height),
        "layouts": layout_ops.list_layout_names(prs),
    }


@mcp.tool()
def create_presentation(
    path: str | None = None,
    template_path: str | None = None,
) -> dict:
    """Create a new PowerPoint presentation and make it the current one.

    Args:
        path: Optional file path where the presentation will be saved later.
        template_path: Optional .pptx or .potx template to base the deck on.
            When omitted, a blank 16:9 presentation is created.
    """
    prs = _state.create(template_path=template_path, target_path=path)

    return {
        "ok": True,
        "message": "Presentation created.",
        "path": str(_state.path) if _state.path else None,
        **_deck_summary(prs),
    }


@mcp.tool()
def open_presentation(path: str) -> dict:
    """Open an existing .pptx file and make it the current presentation.

    Args:
        path: Path to the .pptx file to open.
    """
    prs = _state.load(path)
    return {
        "ok": True,
        "message": "Presentation opened.",
        "path": str(_state.path),
        **_deck_summary(prs),
    }


@mcp.tool()
def save_presentation() -> dict:
    """Save the current presentation to its current path."""
    target = _state.save()
    return {
        "ok": True,
        "message": "Presentation saved.",
        "path": str(target),
    }


@mcp.tool()
def save_as(path: str) -> dict:
    """Save the current presentation to a new path.

    Args:
        path: Destination path for the .pptx file.
    """
    target = _state.save(path)
    return {
        "ok": True,
        "message": "Presentation saved.",
        "path": str(target),
    }


@mcp.tool()
def list_slides() -> dict:
    """List all slides in the current presentation.

    Returns the index, layout name, title, shape count and a short text
    preview for each slide.
    """
    prs = _state.get()
    return {
        "ok": True,
        "slide_count": len(prs.slides),
        "slides": slide_ops.list_slides(prs),
    }


@mcp.tool()
def inspect_slide(index: int) -> dict:
    """Return detailed information about one slide.

    Includes each shape's id, name, type, placeholder status, position and
    size (in inches), plus text and image metadata when present.

    Args:
        index: Zero-based index of the slide.
    """
    prs = _state.get()
    return {"ok": True, **slide_ops.inspect_slide(prs, index)}


@mcp.tool()
def add_slide(
    layout_name: str | None = None,
    index: int | None = None,
) -> dict:
    """Add a new slide to the current presentation.

    Args:
        layout_name: Name of the slide layout to use. Defaults to
            "Title and Content" when available, otherwise "Blank".
        index: Optional zero-based position to insert the slide at. When
            omitted, the slide is appended at the end.
    """
    prs = _state.get()
    slide, final_index = slide_ops.add_slide(
        prs, layout_name=layout_name, index=index
    )
    return {
        "ok": True,
        "message": "Slide added.",
        "index": final_index,
        "slide_count": len(prs.slides),
        "layout": slide.slide_layout.name,
    }


@mcp.tool()
def duplicate_slide(index: int, new_index: int | None = None) -> dict:
    """Duplicate a slide, copying its shapes, background and images.

    Args:
        index: Zero-based index of the slide to duplicate.
        new_index: Optional zero-based position for the copy. When omitted,
            the copy is appended at the end.
    """
    prs = _state.get()
    slide, final_index = slide_ops.duplicate_slide(
        prs, index, new_index=new_index
    )
    return {
        "ok": True,
        "message": "Slide duplicated.",
        "source_index": index,
        "index": final_index,
        "slide_count": len(prs.slides),
        "layout": slide.slide_layout.name,
    }


@mcp.tool()
def delete_slide(index: int) -> dict:
    """Delete a slide from the current presentation.

    Args:
        index: Zero-based index of the slide to delete.
    """
    prs = _state.get()
    slide_ops.delete_slide(prs, index)
    return {
        "ok": True,
        "message": "Slide deleted.",
        "index": index,
        "slide_count": len(prs.slides),
    }


@mcp.tool()
def move_slide(from_index: int, to_index: int) -> dict:
    """Move a slide to a new position.

    Args:
        from_index: Zero-based index of the slide to move.
        to_index: Zero-based destination index.
    """
    prs = _state.get()
    slide_ops.move_slide(prs, from_index, to_index)
    return {
        "ok": True,
        "message": "Slide moved.",
        "from_index": from_index,
        "to_index": to_index,
        "slide_count": len(prs.slides),
    }


@mcp.tool()
def set_slide_text(
    index: int,
    content: dict | str,
    color: str | None = None,
) -> dict:
    """Set the text of a slide.

    Args:
        index: Zero-based index of the slide.
        content: Either a string (sets the slide title) or an object mapping
            targets to text. Targets can be "title", "subtitle", "body",
            "content", a placeholder index as a string, or a shape name.
            Values can be strings or lists of paragraphs.
        color: Optional text color as a hex string (e.g. "#808000") or a
            name such as "darkyellow", "yellow", "red", "blue", "green",
            "orange", "gray", "black" or "white".
    """
    prs = _state.get()
    result = text_ops.set_slide_text(prs, index, content, color=color)
    return {"ok": True, "message": "Slide text updated.", **result}


@mcp.tool()
def add_text_box(
    index: int,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    color: str | None = None,
) -> dict:
    """Add a text box to a slide. Coordinates and size are in inches.

    Args:
        index: Zero-based index of the slide.
        text: Text content (use \\n for multiple paragraphs).
        x: Left position in inches.
        y: Top position in inches.
        w: Width in inches.
        h: Height in inches.
        color: Optional text color as a hex string (e.g. "#808000") or a
            name such as "darkyellow", "yellow", "red", "blue", "green",
            "orange", "gray", "black" or "white".
    """
    prs = _state.get()
    shape = text_ops.add_text_box(prs, index, text, x, y, w, h, color=color)
    return {
        "ok": True,
        "message": "Text box added.",
        "index": index,
        "shape_id": shape.shape_id,
        "name": shape.name,
        "left_in": x,
        "top_in": y,
        "width_in": w,
        "height_in": h,
    }


@mcp.tool()
def replace_text(
    find: str,
    replace: str,
    scope: str | int | None = None,
) -> dict:
    """Replace exact text in the current presentation.

    Args:
        find: Text to search for (must be non-empty).
        replace: Replacement text.
        scope: None or "all" for every slide, or a zero-based slide index.
    """
    prs = _state.get()
    count = text_ops.replace_text(prs, find, replace, scope=scope)
    return {
        "ok": True,
        "message": "Text replaced.",
        "replacements": count,
        "scope": scope if scope is not None else "all",
    }


@mcp.tool()
def add_image(
    index: int,
    source: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
) -> dict:
    """Insert an image from a local path or an http(s) URL.

    Args:
        index: Zero-based index of the slide.
        source: Local file path or http(s) URL.
        x: Left position in inches.
        y: Top position in inches.
        w: Optional width in inches (keeps aspect ratio if h is omitted).
        h: Optional height in inches (keeps aspect ratio if w is omitted).
    """
    prs = _state.get()
    shape = image_ops.add_image(prs, index, source, x, y, w=w, h=h)
    return {
        "ok": True,
        "message": "Image added.",
        "index": index,
        "shape_id": shape.shape_id,
        "name": shape.name,
        "format": shape.image.ext,
        "width_px": shape.image.size[0],
        "height_px": shape.image.size[1],
        "left_in": emu_to_inches(shape.left),
        "top_in": emu_to_inches(shape.top),
        "width_in": emu_to_inches(shape.width),
        "height_in": emu_to_inches(shape.height),
    }


def main() -> None:
    """Run the MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
