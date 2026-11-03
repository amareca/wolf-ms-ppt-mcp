from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from wolf_ms_ppt_mcp.errors import ImageSourceError, UnsupportedImageError
from wolf_ms_ppt_mcp.ops import images as image_ops
from wolf_ms_ppt_mcp.state import PresentationState
from wolf_ms_ppt_mcp.units import emu_to_inches


def _deck(tmp_path: Path) -> Path:
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])  # "Blank"
    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path


def _png(tmp_path: Path, size: tuple[int, int] = (100, 50)) -> Path:
    path = tmp_path / "img.png"
    Image.new("RGB", size, "red").save(path)
    return path


def test_add_image_local(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))
    image_path = _png(tmp_path)

    shape = image_ops.add_image(prs, 0, str(image_path), 1.0, 2.0, w=4.0)

    assert shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    assert emu_to_inches(shape.left) == pytest.approx(1.0)
    assert emu_to_inches(shape.top) == pytest.approx(2.0)
    assert emu_to_inches(shape.width) == pytest.approx(4.0)
    assert emu_to_inches(shape.height) == pytest.approx(2.0)
    assert len(shape.image.blob) > 0


def test_add_image_missing_file_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(FileNotFoundError):
        image_ops.add_image(prs, 0, str(tmp_path / "missing.png"), 0, 0)


def test_add_image_unsupported_scheme_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(UnsupportedImageError):
        image_ops.add_image(prs, 0, "ftp://example.com/img.png", 0, 0)


def test_add_image_from_url(tmp_path: Path, monkeypatch):
    prs = PresentationState().load(_deck(tmp_path))
    png_bytes = _png(tmp_path).read_bytes()

    class FakeResponse:
        content = png_bytes

        def raise_for_status(self):
            return None

    def fake_get(url, **kwargs):
        return FakeResponse()

    monkeypatch.setattr(image_ops.httpx, "get", fake_get)

    shape = image_ops.add_image(
        prs, 0, "https://example.com/img.png", 0.5, 0.5, w=2.0
    )

    assert shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    assert len(shape.image.blob) > 0


def test_add_image_url_error_raises(tmp_path: Path, monkeypatch):
    prs = PresentationState().load(_deck(tmp_path))

    def fake_get(url, **kwargs):
        raise RuntimeError("network down")

    monkeypatch.setattr(image_ops.httpx, "get", fake_get)

    with pytest.raises(ImageSourceError):
        image_ops.add_image(prs, 0, "https://example.com/img.png", 0, 0)


def test_add_image_persists_after_save(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))
    image_path = _png(tmp_path)

    image_ops.add_image(prs, 0, str(image_path), 1.0, 1.0, w=2.0)
    output = tmp_path / "output.pptx"
    prs.save(str(output))

    reopened = Presentation(str(output))
    pictures = [
        shape
        for shape in reopened.slides[0].shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]

    assert len(pictures) == 1
    assert len(pictures[0].image.blob) > 0
