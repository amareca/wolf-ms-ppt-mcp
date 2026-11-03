"""Text operations: set slide text, add text boxes and find/replace."""

from __future__ import annotations

from typing import Any

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.shapes.base import BaseShape
from pptx.slide import Slide

from ..errors import TextTargetNotFoundError, WolfPptError
from ..units import inches_to_emu
from ._common import validate_slide_index


def _text_shapes(slide: Slide) -> list[BaseShape]:
    return [shape for shape in slide.shapes if shape.has_text_frame]


def _find_placeholder(slide: Slide, placeholder_type) -> BaseShape | None:
    for shape in slide.shapes:
        if not shape.is_placeholder:
            continue
        if shape.placeholder_format.type == placeholder_type:
            return shape
    return None


def _resolve_text_target(slide: Slide, selector: str) -> BaseShape:
    """Resolve a selector to a text-bearing shape on the slide.

    Supported selectors:
    - ``"title"``: the slide title placeholder.
    - ``"subtitle"``: the subtitle placeholder.
    - ``"body"`` / ``"content"``: the first body/content placeholder.
    - a digit string: a placeholder by its placeholder index.
    - any other string: a shape name (case-insensitive fallback).
    """
    key = selector.strip()
    lowered = key.lower()

    if lowered == "title":
        shape = slide.shapes.title
        if shape is None:
            raise TextTargetNotFoundError(
                "This slide has no title placeholder. "
                "Use add_text_box or a shape name instead."
            )
        return shape

    if lowered == "subtitle":
        shape = _find_placeholder(slide, PP_PLACEHOLDER.SUBTITLE)
        if shape is None:
            raise TextTargetNotFoundError("This slide has no subtitle placeholder.")
        return shape

    if lowered in ("body", "content"):
        for placeholder_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
            shape = _find_placeholder(slide, placeholder_type)
            if shape is not None:
                return shape
        raise TextTargetNotFoundError(
            "This slide has no body/content placeholder. "
            "Use add_text_box or a shape name instead."
        )

    if key.isdigit():
        try:
            return slide.placeholders[int(key)]
        except KeyError as exc:
            raise TextTargetNotFoundError(
                f"No placeholder with index {key} exists on this slide."
            ) from exc

    for shape in slide.shapes:
        if shape.name == key:
            return shape
    for shape in slide.shapes:
        if shape.name.lower() == lowered:
            return shape

    available = ", ".join(shape.name for shape in _text_shapes(slide)) or "(none)"
    raise TextTargetNotFoundError(
        f"Text target '{selector}' not found. "
        f"Text shapes on this slide: {available}."
    )


def _set_shape_text(shape: BaseShape, value: str | list[str]) -> None:
    if not shape.has_text_frame:
        raise WolfPptError(f"Shape '{shape.name}' has no text frame.")
    if isinstance(value, (list, tuple)):
        text = "\n".join(str(item) for item in value)
    else:
        text = str(value)
    shape.text_frame.text = text


def set_slide_text(
    prs: Presentation,
    index: int,
    content: str | dict[str, Any],
) -> dict[str, Any]:
    """Set text on a slide.

    - A string sets the slide title.
    - A dict maps selectors (``title``, ``subtitle``, ``body``, a placeholder
      index, or a shape name) to strings or lists of paragraphs.
    """
    validate_slide_index(prs, index)
    slide = prs.slides[index]

    if isinstance(content, str):
        shape = _resolve_text_target(slide, "title")
        _set_shape_text(shape, content)
        return {"index": index, "updated": ["title"]}

    if not isinstance(content, dict):
        raise WolfPptError(
            "content must be a string or an object mapping targets to text."
        )

    # Resolve every target before mutating anything, so a bad selector does
    # not leave the slide partially updated.
    resolved = [
        (str(selector), _resolve_text_target(slide, str(selector)), value)
        for selector, value in content.items()
    ]
    for _selector, shape, value in resolved:
        _set_shape_text(shape, value)

    return {"index": index, "updated": [selector for selector, _, _ in resolved]}


def add_text_box(
    prs: Presentation,
    index: int,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
) -> BaseShape:
    """Add a text box to a slide. Coordinates and size are in inches."""
    validate_slide_index(prs, index)
    if w <= 0 or h <= 0:
        raise WolfPptError("Text box width and height must be positive.")

    slide = prs.slides[index]
    shape = slide.shapes.add_textbox(
        inches_to_emu(x),
        inches_to_emu(y),
        inches_to_emu(w),
        inches_to_emu(h),
    )
    shape.text_frame.text = text
    return shape


def _slides_for_scope(prs: Presentation, scope: str | int | None) -> list[Slide]:
    if scope is None:
        return list(prs.slides)
    if isinstance(scope, int):
        validate_slide_index(prs, scope)
        return [prs.slides[scope]]
    if isinstance(scope, str):
        normalized = scope.strip().lower()
        if normalized in ("", "all"):
            return list(prs.slides)
        if normalized.isdigit():
            index = int(normalized)
            validate_slide_index(prs, index)
            return [prs.slides[index]]
    raise WolfPptError("scope must be None, 'all', or a zero-based slide index.")


def replace_text(
    prs: Presentation,
    find: str,
    replace: str,
    scope: str | int | None = None,
) -> int:
    """Replace exact text in runs and return the number of replacements.

    Replacement is performed run by run, so a match split across differently
    formatted runs will not be replaced.
    """
    if not find:
        raise WolfPptError("find must be a non-empty string.")

    count = 0
    for slide in _slides_for_scope(prs, scope):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if find in run.text:
                        count += run.text.count(find)
                        run.text = run.text.replace(find, replace)
    return count
