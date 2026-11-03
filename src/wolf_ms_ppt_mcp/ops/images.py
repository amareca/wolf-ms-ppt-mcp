"""Image insertion operations."""

from __future__ import annotations

import io
from pathlib import Path

import httpx
from pptx import Presentation
from pptx.shapes.base import BaseShape

from ..errors import ImageSourceError, UnsupportedImageError, WolfPptError
from ..units import inches_to_emu
from ._common import validate_slide_index

_MAX_IMAGE_BYTES = 20 * 1024 * 1024


def _is_url(source: str) -> bool:
    return source.lower().startswith(("http://", "https://"))


def _download_image(source: str) -> io.BytesIO:
    """Download an image into memory and return it as a byte stream."""
    try:
        response = httpx.get(source, follow_redirects=True, timeout=30.0)
        response.raise_for_status()
    except Exception as exc:
        raise ImageSourceError(
            f"Could not download image from '{source}': {exc}"
        ) from exc

    content = response.content
    if len(content) > _MAX_IMAGE_BYTES:
        limit_mb = _MAX_IMAGE_BYTES // (1024 * 1024)
        raise ImageSourceError(
            f"Image at '{source}' exceeds the {limit_mb} MB limit."
        )
    return io.BytesIO(content)


def add_image(
    prs: Presentation,
    index: int,
    source: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
) -> BaseShape:
    """Insert an image from a local path or an http(s) URL.

    Coordinates and sizes are in inches. When only one of ``w``/``h`` is
    provided, the image keeps its aspect ratio.
    """
    validate_slide_index(prs, index)
    if w is not None and w <= 0:
        raise WolfPptError("Image width must be positive.")
    if h is not None and h <= 0:
        raise WolfPptError("Image height must be positive.")

    if _is_url(source):
        image_file: str | io.BytesIO = _download_image(source)
    else:
        if "://" in source:
            raise UnsupportedImageError(
                f"Unsupported image source '{source}'. "
                "Use a local path or an http(s) URL."
            )
        path = Path(source).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"Image not found: {path}")
        image_file = str(path)

    slide = prs.slides[index]
    try:
        shape = slide.shapes.add_picture(
            image_file,
            inches_to_emu(x),
            inches_to_emu(y),
            width=inches_to_emu(w) if w is not None else None,
            height=inches_to_emu(h) if h is not None else None,
        )
    except Exception as exc:
        raise UnsupportedImageError(
            f"Could not add image from '{source}'. "
            "It may not be a supported image format."
        ) from exc

    return shape
