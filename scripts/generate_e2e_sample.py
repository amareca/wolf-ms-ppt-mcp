"""Generate a sample deck for the Phase 3 end-to-end test.

All content operations go through the project's own operations:

- Phase 2: add_slide, duplicate_slide, move_slide, delete_slide.
- Phase 3: set_slide_text, add_text_box, replace_text, add_image.

Only the source PNG is created with Pillow.

Run from the project root:
    uv run python scripts/generate_e2e_sample.py
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from wolf_ms_ppt_mcp.ops import images as image_ops
from wolf_ms_ppt_mcp.ops import slides as slide_ops
from wolf_ms_ppt_mcp.ops import text as text_ops
from wolf_ms_ppt_mcp.state import PresentationState

SAMPLES_DIR = Path("samples")
OUTPUT_PATH = SAMPLES_DIR / "fase3-e2e.pptx"
IMAGE_PATH = SAMPLES_DIR / "e2e-image.png"


def _make_image() -> None:
    image = Image.new("RGB", (600, 300), "#1f4e79")
    draw = ImageDraw.Draw(image)
    draw.text((40, 130), "IMAGEN E2E", fill="white")
    image.save(IMAGE_PATH)


def _build_deck() -> None:
    state = PresentationState()
    prs = state.create(target_path=OUTPUT_PATH)

    # 0: title slide
    _, index = slide_ops.add_slide(prs, layout_name="Title Slide")
    text_ops.set_slide_text(
        prs,
        index,
        {
            "title": "FASE 3 — PRUEBA E2E",
            "subtitle": "Generado con wolf-ms-ppt-mcp",
        },
    )

    # 1: "Diapositiva A"
    _, index = slide_ops.add_slide(prs, layout_name="Title and Content")
    text_ops.set_slide_text(
        prs, index, {"title": "Diapositiva A", "body": "Contenido de A"}
    )

    # 2: "Diapositiva B" with a two-paragraph body
    _, index = slide_ops.add_slide(prs, layout_name="Title and Content")
    text_ops.set_slide_text(
        prs,
        index,
        {
            "title": "Diapositiva B",
            "body": ["Contenido de B", "Segunda línea de B"],
        },
    )

    # 3: blank slide with an image and a text box
    _, index = slide_ops.add_slide(prs, layout_name="Blank")
    image_ops.add_image(prs, index, str(IMAGE_PATH), x=3.5, y=2.0, w=6.0)
    text_ops.add_text_box(
        prs,
        index,
        "Imagen insertada con add_image",
        x=3.5,
        y=5.5,
        w=6.0,
        h=1.0,
    )

    # Exercise replace_text on the "Diapositiva B" slide (index 2).
    text_ops.replace_text(
        prs, "Contenido de B", "Contenido actualizado de B", scope=2
    )

    # Phase 2 operations:
    # [Title, A, B, Image] -> [Title, A, B, Image, ImageCopy]
    slide_ops.duplicate_slide(prs, 3)
    # -> [Title, ImageCopy, A, B, Image]
    slide_ops.move_slide(prs, 4, 1)
    # -> [Title, ImageCopy, Blank, A, B, Image]
    slide_ops.add_slide(prs, layout_name="Blank", index=2)
    # delete "A" -> [Title, ImageCopy, Blank, B, Image]
    slide_ops.delete_slide(prs, 3)

    state.save()


def _verify() -> None:
    prs = Presentation(str(OUTPUT_PATH))
    print(f"file: {OUTPUT_PATH.resolve()}")
    print(f"slides: {len(prs.slides)}")
    for summary in slide_ops.list_slides(prs):
        slide = prs.slides[summary["index"]]
        pictures = [
            shape
            for shape in slide.shapes
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        ]
        for picture in pictures:
            assert len(picture.image.blob) > 0
        print(
            f"  {summary['index']} | {summary['layout']} | "
            f"title={summary['title']!r} | pictures={len(pictures)} | "
            f"preview={summary['text_preview']!r}"
        )


def main() -> None:
    SAMPLES_DIR.mkdir(exist_ok=True)
    _make_image()
    _build_deck()
    _verify()


if __name__ == "__main__":
    main()
