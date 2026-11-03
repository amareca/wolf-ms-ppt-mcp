"""In-memory state for the single "current" presentation.

The model mirrors PowerPoint's ``ActivePresentation``: the agent opens or
creates a deck once, performs operations on it, and saves when done.
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from .errors import PresentationNotOpenError, SaveError

DEFAULT_WIDTH_IN = 13.333
DEFAULT_HEIGHT_IN = 7.5


class PresentationState:
    """Holds the current presentation and its associated file path."""

    def __init__(self) -> None:
        self._prs: Presentation | None = None
        self._path: Path | None = None

    @property
    def is_open(self) -> bool:
        return self._prs is not None

    @property
    def path(self) -> Path | None:
        return self._path

    def load(self, path: str | Path) -> Presentation:
        """Load an existing ``.pptx`` file as the current presentation."""
        resolved = Path(path).expanduser().resolve()
        if not resolved.exists():
            raise FileNotFoundError(f"Presentation not found: {resolved}")
        self._prs = Presentation(str(resolved))
        self._path = resolved
        return self._prs

    def create(
        self,
        template_path: str | Path | None = None,
        target_path: str | Path | None = None,
    ) -> Presentation:
        """Create a new presentation, optionally from a template.

        ``target_path`` is stored as the default location used by
        :meth:`save` when no explicit path is given.
        """
        if template_path is not None:
            resolved = Path(template_path).expanduser().resolve()
            if not resolved.exists():
                raise FileNotFoundError(f"Template not found: {resolved}")
            self._prs = Presentation(str(resolved))
        else:
            self._prs = Presentation()
            self._prs.slide_width = Inches(DEFAULT_WIDTH_IN)
            self._prs.slide_height = Inches(DEFAULT_HEIGHT_IN)

        self._path = (
            Path(target_path).expanduser().resolve()
            if target_path is not None
            else None
        )
        return self._prs

    def get(self) -> Presentation:
        """Return the current presentation or raise if none is open."""
        if self._prs is None:
            raise PresentationNotOpenError(
                "No presentation is open. Use create_presentation or "
                "open_presentation first."
            )
        return self._prs

    def save(self, path: str | Path | None = None) -> Path:
        """Save the current presentation and return the written path."""
        prs = self.get()

        if path is not None:
            target = Path(path).expanduser().resolve()
        elif self._path is not None:
            target = self._path
        else:
            raise SaveError(
                "No save path is set. Call create_presentation(path=...) or "
                "use save_as(path=...)."
            )

        target.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(target))
        self._path = target
        return target
