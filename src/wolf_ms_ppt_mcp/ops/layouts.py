"""Slide layout resolution helpers."""

from __future__ import annotations

from pptx import Presentation
from pptx.slide import SlideLayout

from ..errors import LayoutNotFoundError

DEFAULT_LAYOUT_CANDIDATES = ("Title and Content", "Blank")


def list_layout_names(prs: Presentation) -> list[str]:
    """Return the names of all layouts available in the presentation."""
    return [layout.name for layout in prs.slide_layouts]


def resolve_layout(prs: Presentation, layout_name: str | None) -> SlideLayout:
    """Return a layout by name (case-insensitive), with sensible defaults.

    When ``layout_name`` is None, "Title and Content" is preferred, then
    "Blank", then the first available layout.
    """
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise LayoutNotFoundError("The presentation has no slide layouts.")

    if layout_name is None:
        for candidate in DEFAULT_LAYOUT_CANDIDATES:
            for layout in layouts:
                if layout.name == candidate:
                    return layout
        return layouts[0]

    for layout in layouts:
        if layout.name == layout_name:
            return layout

    lowered = layout_name.strip().lower()
    for layout in layouts:
        if layout.name.lower() == lowered:
            return layout

    available = ", ".join(layout.name for layout in layouts)
    raise LayoutNotFoundError(
        f"Layout '{layout_name}' not found. Available layouts: {available}."
    )
