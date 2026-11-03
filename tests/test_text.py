from pathlib import Path

import pytest
from pptx import Presentation
from pptx.dml.color import RGBColor

from wolf_ms_ppt_mcp.errors import (
    SlideIndexError,
    TextTargetNotFoundError,
    WolfPptError,
)
from wolf_ms_ppt_mcp.ops import text as text_ops
from wolf_ms_ppt_mcp.state import PresentationState
from wolf_ms_ppt_mcp.units import emu_to_inches


def _deck(tmp_path: Path) -> Path:
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide
    slide.shapes.title.text = "Old title"
    slide.placeholders[1].text = "Old subtitle"

    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
    slide.shapes.title.text = "Second"
    slide.placeholders[1].text = "Body text"

    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path


def test_set_slide_text_string_sets_title(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    result = text_ops.set_slide_text(prs, 0, "New title")

    assert result["updated"] == ["title"]
    assert prs.slides[0].shapes.title.text == "New title"


def test_set_slide_text_dict_sets_title_and_body(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    result = text_ops.set_slide_text(
        prs, 1, {"title": "New title", "body": "New body"}
    )

    assert result["updated"] == ["title", "body"]
    assert prs.slides[1].shapes.title.text == "New title"
    assert prs.slides[1].placeholders[1].text == "New body"


def test_set_slide_text_by_shape_name(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))
    shape_name = prs.slides[1].shapes.title.name

    text_ops.set_slide_text(prs, 1, {shape_name: "By name"})

    assert prs.slides[1].shapes.title.text == "By name"


def test_set_slide_text_missing_target_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(TextTargetNotFoundError):
        text_ops.set_slide_text(prs, 1, {"does-not-exist": "x"})


def test_set_slide_text_multiline_creates_paragraphs(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    text_ops.set_slide_text(prs, 1, {"body": "Line 1\nLine 2"})

    paragraphs = prs.slides[1].placeholders[1].text_frame.paragraphs
    assert [p.text for p in paragraphs] == ["Line 1", "Line 2"]


def test_set_slide_text_out_of_range_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(SlideIndexError):
        text_ops.set_slide_text(prs, 9, "x")


def test_add_text_box(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    shape = text_ops.add_text_box(prs, 1, "Hello", 1.0, 2.0, 3.0, 1.5)

    assert shape.text_frame.text == "Hello"
    assert emu_to_inches(shape.left) == pytest.approx(1.0)
    assert emu_to_inches(shape.top) == pytest.approx(2.0)
    assert emu_to_inches(shape.width) == pytest.approx(3.0)
    assert emu_to_inches(shape.height) == pytest.approx(1.5)


def test_add_text_box_invalid_size_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(WolfPptError):
        text_ops.add_text_box(prs, 1, "Hello", 0, 0, 0, 1)


def test_replace_text_all_slides(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    count = text_ops.replace_text(prs, "Old", "New")

    assert count == 2
    assert prs.slides[0].shapes.title.text == "New title"
    assert prs.slides[0].placeholders[1].text == "New subtitle"


def test_replace_text_single_slide(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    count = text_ops.replace_text(prs, "Body", "Content", scope=1)

    assert count == 1
    assert prs.slides[1].placeholders[1].text == "Content text"


def test_replace_text_accepts_numeric_string_scope(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    count = text_ops.replace_text(prs, "Body", "Content", scope="1")

    assert count == 1


def test_replace_text_empty_find_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(WolfPptError):
        text_ops.replace_text(prs, "", "x")


def test_replace_text_invalid_scope_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(WolfPptError):
        text_ops.replace_text(prs, "Old", "New", scope="nope")


def test_set_slide_text_color_hex(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    text_ops.set_slide_text(prs, 1, {"body": "Colored"}, color="#808000")

    run = prs.slides[1].placeholders[1].text_frame.paragraphs[0].runs[0]
    assert run.font.color.rgb == RGBColor(0x80, 0x80, 0x00)


def test_set_slide_text_color_named(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    text_ops.set_slide_text(prs, 1, {"title": "Colored"}, color="darkyellow")

    run = prs.slides[1].shapes.title.text_frame.paragraphs[0].runs[0]
    assert run.font.color.rgb == RGBColor(0x80, 0x80, 0x00)


def test_add_text_box_color(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    shape = text_ops.add_text_box(
        prs, 1, "Box", 1, 1, 3, 1, color="yellow"
    )

    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.font.color.rgb == RGBColor(0xFF, 0xFF, 0x00)


def test_invalid_color_raises(tmp_path: Path):
    prs = PresentationState().load(_deck(tmp_path))

    with pytest.raises(WolfPptError):
        text_ops.set_slide_text(prs, 1, {"body": "x"}, color="notacolor")
