from pathlib import Path

import pytest

from wolf_ms_ppt_mcp.errors import (
    PresentationNotOpenError,
    SaveError,
    SlideIndexError,
)
from wolf_ms_ppt_mcp.ops import slides as slide_ops
from wolf_ms_ppt_mcp.state import PresentationState
from wolf_ms_ppt_mcp.units import emu_to_inches, inches_to_emu


def test_create_sets_16_9():
    state = PresentationState()
    prs = state.create()

    assert emu_to_inches(prs.slide_width) == pytest.approx(13.333, abs=0.01)
    assert emu_to_inches(prs.slide_height) == pytest.approx(7.5, abs=0.01)
    assert len(prs.slides) == 0


def test_get_without_open_raises():
    state = PresentationState()
    with pytest.raises(PresentationNotOpenError):
        state.get()


def test_save_and_reopen(tmp_path: Path):
    state = PresentationState()
    state.create(target_path=tmp_path / "deck.pptx")
    saved = state.save()

    assert saved.exists()
    assert saved.suffix == ".pptx"

    reopened = PresentationState()
    prs = reopened.load(saved)
    assert len(prs.slides) == 0


def test_save_without_path_raises():
    state = PresentationState()
    state.create()
    with pytest.raises(SaveError):
        state.save()


def test_load_missing_file_raises(tmp_path: Path):
    state = PresentationState()
    with pytest.raises(FileNotFoundError):
        state.load(tmp_path / "missing.pptx")


def test_list_slides(sample_deck: Path):
    state = PresentationState()
    prs = state.load(sample_deck)

    slides = slide_ops.list_slides(prs)

    assert len(slides) == 1
    assert slides[0]["index"] == 0
    assert slides[0]["title"] == "Hello World"
    assert "Hello World" in slides[0]["text_preview"]
    assert slides[0]["shape_count"] >= 1


def test_inspect_slide(sample_deck: Path):
    state = PresentationState()
    prs = state.load(sample_deck)

    info = slide_ops.inspect_slide(prs, 0)

    assert info["index"] == 0
    assert info["shape_count"] >= 1
    assert any("Hello World" in (shape.get("text") or "") for shape in info["shapes"])
    assert all("shape_type" in shape for shape in info["shapes"])


def test_inspect_slide_out_of_range(sample_deck: Path):
    state = PresentationState()
    prs = state.load(sample_deck)

    with pytest.raises(SlideIndexError):
        slide_ops.inspect_slide(prs, 5)

    with pytest.raises(SlideIndexError):
        slide_ops.inspect_slide(prs, -1)


def test_units_roundtrip():
    assert emu_to_inches(inches_to_emu(2.5)) == pytest.approx(2.5)
    assert emu_to_inches(inches_to_emu(13.333)) == pytest.approx(13.333, abs=0.001)
