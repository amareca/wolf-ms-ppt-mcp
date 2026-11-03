from pathlib import Path

import pytest
from pptx import Presentation


@pytest.fixture
def sample_deck(tmp_path: Path) -> Path:
    """Create a small .pptx with one titled slide and return its path."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Hello World"
    path = tmp_path / "sample.pptx"
    prs.save(str(path))
    return path
