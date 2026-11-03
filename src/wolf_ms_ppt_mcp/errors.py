"""Domain errors raised by the MCP server.

These errors carry user-facing messages that are safe to return to the AI
agent through the MCP protocol.
"""


class WolfPptError(Exception):
    """Base class for all domain errors raised by the server."""


class PresentationNotOpenError(WolfPptError):
    """Raised when an operation requires an open presentation."""


class SlideIndexError(WolfPptError):
    """Raised when a slide index is out of range."""


class LayoutNotFoundError(WolfPptError):
    """Raised when a requested slide layout does not exist."""


class SaveError(WolfPptError):
    """Raised when the presentation cannot be saved."""


class UnsupportedImageError(WolfPptError):
    """Raised when an image source is not supported."""


class TextTargetNotFoundError(WolfPptError):
    """Raised when a requested text target cannot be found on a slide."""


class ImageSourceError(WolfPptError):
    """Raised when an image cannot be loaded from its source."""
