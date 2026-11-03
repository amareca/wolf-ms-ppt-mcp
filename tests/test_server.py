from pathlib import Path

import pytest

from wolf_ms_ppt_mcp import server
from wolf_ms_ppt_mcp.errors import PresentationNotOpenError
from wolf_ms_ppt_mcp.state import PresentationState


@pytest.fixture(autouse=True)
def fresh_state():
    """Reset the server's global presentation state before each test."""
    server._state = PresentationState()
    yield


def test_create_save_open_and_list(tmp_path: Path):
    target = tmp_path / "deck.pptx"

    created = server.create_presentation(path=str(target))
    assert created["ok"] is True
    assert created["slide_count"] == 0
    assert created["slide_width_in"] == pytest.approx(13.333, abs=0.01)
    assert created["slide_height_in"] == pytest.approx(7.5, abs=0.01)

    saved = server.save_presentation()
    assert Path(saved["path"]).exists()

    opened = server.open_presentation(str(target))
    assert opened["ok"] is True
    assert opened["slide_count"] == 0

    listed = server.list_slides()
    assert listed["slide_count"] == 0
    assert listed["slides"] == []


def test_inspect_slide_requires_open_presentation():
    with pytest.raises(PresentationNotOpenError):
        server.inspect_slide(0)


def test_save_without_open_presentation_raises():
    with pytest.raises(PresentationNotOpenError):
        server.save_presentation()


def test_slide_mutation_tools():
    server.create_presentation()

    first = server.add_slide()
    assert first["ok"] is True
    assert first["slide_count"] == 1
    assert first["index"] == 0

    second = server.add_slide()
    assert second["index"] == 1

    duplicated = server.duplicate_slide(0)
    assert duplicated["slide_count"] == 3
    assert duplicated["index"] == 2

    moved = server.move_slide(0, 2)
    assert moved["ok"] is True
    assert moved["slide_count"] == 3

    deleted = server.delete_slide(0)
    assert deleted["ok"] is True
    assert deleted["slide_count"] == 2


def test_text_and_image_tools(tmp_path: Path):
    from PIL import Image

    image_path = tmp_path / "img.png"
    Image.new("RGB", (20, 10), "blue").save(image_path)

    server.create_presentation()
    server.add_slide()

    updated = server.set_slide_text(0, {"title": "Hello"})
    assert updated["ok"] is True
    assert updated["updated"] == ["title"]

    box = server.add_text_box(0, "Box text", 1, 1, 3, 1)
    assert box["ok"] is True
    assert box["name"]

    replaced = server.replace_text("Hello", "Hi")
    assert replaced["replacements"] == 1

    added = server.add_image(0, str(image_path), 1, 2, w=2.0)
    assert added["ok"] is True
    assert added["format"] == "png"
