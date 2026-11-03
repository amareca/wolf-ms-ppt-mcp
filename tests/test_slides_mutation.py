from pathlib import Path
import warnings

import pytest
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

from wolf_ms_ppt_mcp.errors import LayoutNotFoundError, SlideIndexError
from wolf_ms_ppt_mcp.ops import slides as slide_ops
from wolf_ms_ppt_mcp.state import PresentationState


def _deck_with_titles(tmp_path: Path, titles: list[str]) -> Path:
    prs = Presentation()
    for title in titles:
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # "Title Only"
        slide.shapes.title.text = title
    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path


def _titles(prs: Presentation) -> list[str]:
    return [slide.shapes.title.text for slide in prs.slides]


# --------------------------------------------------------------------------- #
# add_slide
# --------------------------------------------------------------------------- #


def test_add_slide_default_layout():
    prs = Presentation()

    slide, index = slide_ops.add_slide(prs)

    assert index == 0
    assert len(prs.slides) == 1
    assert slide.slide_layout.name == "Title and Content"


def test_add_slide_at_index(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B"]))

    slide, index = slide_ops.add_slide(prs, layout_name="Blank", index=1)

    assert index == 1
    assert len(prs.slides) == 3
    assert prs.slides[1].slide_id == slide.slide_id
    assert prs.slides[1].slide_layout.name == "Blank"


def test_add_slide_appends_when_index_equals_count(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A"]))

    slide, index = slide_ops.add_slide(prs, index=1)

    assert index == 1
    assert len(prs.slides) == 2


def test_add_slide_unknown_layout_raises():
    prs = Presentation()

    with pytest.raises(LayoutNotFoundError):
        slide_ops.add_slide(prs, layout_name="Nope")


def test_add_slide_index_out_of_range_raises(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A"]))

    with pytest.raises(SlideIndexError):
        slide_ops.add_slide(prs, index=5)


# --------------------------------------------------------------------------- #
# move_slide
# --------------------------------------------------------------------------- #


def test_move_slide_forward(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B", "C"]))

    slide_ops.move_slide(prs, 0, 2)

    assert _titles(prs) == ["B", "C", "A"]


def test_move_slide_backward(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B", "C"]))

    slide_ops.move_slide(prs, 2, 0)

    assert _titles(prs) == ["C", "A", "B"]


def test_move_slide_to_same_index_is_noop(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B"]))

    slide_ops.move_slide(prs, 1, 1)

    assert _titles(prs) == ["A", "B"]


def test_move_slide_out_of_range_raises(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B"]))

    with pytest.raises(SlideIndexError):
        slide_ops.move_slide(prs, 0, 5)

    with pytest.raises(SlideIndexError):
        slide_ops.move_slide(prs, 5, 0)


# --------------------------------------------------------------------------- #
# delete_slide
# --------------------------------------------------------------------------- #


def test_delete_slide(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B", "C"]))

    slide_ops.delete_slide(prs, 1)

    assert _titles(prs) == ["A", "C"]
    assert len(prs.slides) == 2


def test_delete_slide_out_of_range_raises(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A"]))

    with pytest.raises(SlideIndexError):
        slide_ops.delete_slide(prs, 3)


# --------------------------------------------------------------------------- #
# duplicate_slide
# --------------------------------------------------------------------------- #


def test_duplicate_slide_copies_text(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B"]))

    slide, index = slide_ops.duplicate_slide(prs, 0)

    assert index == 2
    assert len(prs.slides) == 3
    assert prs.slides[2].slide_id == slide.slide_id
    assert prs.slides[2].shapes.title.text == "A"


def test_duplicate_slide_at_index(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B"]))

    slide, index = slide_ops.duplicate_slide(prs, 0, new_index=1)

    assert index == 1
    assert _titles(prs) == ["A", "A", "B"]


def test_duplicate_slide_copies_image(tmp_path: Path):
    image_path = tmp_path / "red.png"
    Image.new("RGB", (10, 10), "red").save(image_path)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # "Blank"
    slide.shapes.add_picture(
        str(image_path), Inches(1), Inches(1), Inches(1), Inches(1)
    )
    deck_path = tmp_path / "deck.pptx"
    prs.save(str(deck_path))

    prs = PresentationState().load(deck_path)
    slide_ops.duplicate_slide(prs, 0)

    duplicated = prs.slides[1]
    pictures = [
        shape
        for shape in duplicated.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]

    assert len(pictures) == 1
    # Accessing the image blob validates that the relationship was remapped.
    assert len(pictures[0].image.blob) > 0


def test_duplicate_slide_index_out_of_range_raises(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A"]))

    with pytest.raises(SlideIndexError):
        slide_ops.duplicate_slide(prs, 3)


def test_mutations_persist_after_save(tmp_path: Path):
    prs = PresentationState().load(_deck_with_titles(tmp_path, ["A", "B", "C"]))

    slide_ops.move_slide(prs, 0, 2)  # [B, C, A]
    slide_ops.delete_slide(prs, 0)  # [C, A]
    slide_ops.duplicate_slide(prs, 0)  # [C, A, C]

    output = tmp_path / "output.pptx"
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        prs.save(str(output))

    duplicate_warnings = [
        item for item in caught if "Duplicate name" in str(item.message)
    ]
    assert duplicate_warnings == []

    reopened = Presentation(str(output))
    assert [slide.shapes.title.text for slide in reopened.slides] == [
        "C",
        "A",
        "C",
    ]


def test_duplicated_image_persists_after_save(tmp_path: Path):
    image_path = tmp_path / "red.png"
    Image.new("RGB", (10, 10), "red").save(image_path)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # "Blank"
    slide.shapes.add_picture(
        str(image_path), Inches(1), Inches(1), Inches(1), Inches(1)
    )
    deck_path = tmp_path / "deck.pptx"
    prs.save(str(deck_path))

    prs = PresentationState().load(deck_path)
    slide_ops.duplicate_slide(prs, 0)
    output = tmp_path / "output.pptx"
    prs.save(str(output))

    reopened = Presentation(str(output))
    pictures = [
        shape
        for shape in reopened.slides[1].shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    assert len(pictures) == 1
    assert len(pictures[0].image.blob) > 0
