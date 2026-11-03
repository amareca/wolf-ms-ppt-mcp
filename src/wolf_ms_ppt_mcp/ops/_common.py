"""Shared helpers for slide operations."""

from __future__ import annotations

from pptx import Presentation

from ..errors import SlideIndexError


def validate_slide_index(prs: Presentation, index: int) -> None:
    """Raise ``SlideIndexError`` when ``index`` is not a valid slide index."""
    count = len(prs.slides)
    if index < 0 or index >= count:
        raise SlideIndexError(
            f"Slide index {index} is out of range (0..{count - 1}); "
            f"the presentation has {count} slide(s)."
        )
