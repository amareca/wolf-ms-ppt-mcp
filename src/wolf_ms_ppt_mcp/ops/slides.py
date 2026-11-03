"""Slide operations: listing, inspection and mutation."""

from __future__ import annotations

import copy
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.slide import Slide

from ..errors import SlideIndexError
from ..units import emu_to_inches
from ._common import validate_slide_index as _validate_index
from .layouts import resolve_layout

# Relationship attributes that can appear inside a slide's XML.
_REL_ATTRS = ("r:id", "r:embed", "r:link", "r:pict", "r:dm", "r:lo", "r:qs", "r:cs")

# Relationships that must not be copied when duplicating a slide: the new
# slide already has its own layout, and notes are out of scope for the MVP.
_SKIP_RELTYPES = frozenset({RT.SLIDE_LAYOUT, RT.NOTES_SLIDE})


# --------------------------------------------------------------------------- #
# Read operations
# --------------------------------------------------------------------------- #


def _slide_title(slide) -> str | None:
    """Return the slide title text, or None when there is no title."""
    title = slide.shapes.title
    if title is None or not title.has_text_frame:
        return None
    text = title.text_frame.text.strip()
    return text or None


def _text_preview(slide, limit: int = 120) -> str:
    """Build a short single-line preview from all text frames on a slide."""
    parts: list[str] = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text:
                parts.append(text)
    preview = " | ".join(parts)
    if len(preview) > limit:
        preview = preview[: limit - 1].rstrip() + "\u2026"
    return preview


def list_slides(prs: Presentation) -> list[dict[str, Any]]:
    """Return a compact summary for every slide in the presentation."""
    result: list[dict[str, Any]] = []
    for index, slide in enumerate(prs.slides):
        result.append(
            {
                "index": index,
                "layout": slide.slide_layout.name,
                "title": _slide_title(slide),
                "shape_count": len(slide.shapes),
                "text_preview": _text_preview(slide),
            }
        )
    return result


def _shape_type_name(shape) -> str:
    try:
        return str(shape.shape_type)
    except Exception:  # pragma: no cover - defensive
        return "UNKNOWN"


def _shape_info(shape) -> dict[str, Any]:
    """Serialize a single shape into a JSON-friendly dictionary."""
    info: dict[str, Any] = {
        "shape_id": shape.shape_id,
        "name": shape.name,
        "shape_type": _shape_type_name(shape),
        "is_placeholder": bool(shape.is_placeholder),
        "left_in": emu_to_inches(shape.left) if shape.left is not None else None,
        "top_in": emu_to_inches(shape.top) if shape.top is not None else None,
        "width_in": emu_to_inches(shape.width) if shape.width is not None else None,
        "height_in": emu_to_inches(shape.height) if shape.height is not None else None,
    }

    if shape.is_placeholder:
        try:
            info["placeholder_type"] = str(shape.placeholder_format.type)
        except Exception:  # pragma: no cover - defensive
            info["placeholder_type"] = None

    if shape.has_text_frame:
        info["text"] = shape.text_frame.text

    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        try:
            image = shape.image
            info["image"] = {
                "filename": image.filename,
                "format": image.ext,
                "width_px": image.size[0],
                "height_px": image.size[1],
            }
        except Exception:  # pragma: no cover - defensive
            pass

    return info


def inspect_slide(prs: Presentation, index: int) -> dict[str, Any]:
    """Return detailed information about one slide."""
    _validate_index(prs, index)
    slide = prs.slides[index]
    return {
        "index": index,
        "layout": slide.slide_layout.name,
        "title": _slide_title(slide),
        "shape_count": len(slide.shapes),
        "shapes": [_shape_info(shape) for shape in slide.shapes],
    }


# --------------------------------------------------------------------------- #
# Mutation operations
# --------------------------------------------------------------------------- #


def _sld_id_elements(prs: Presentation) -> list:
    """Return the ``<p:sldId>`` elements in presentation order."""
    return list(prs.slides._sldIdLst.sldId_lst)


def _normalize_slide_partnames(prs: Presentation) -> None:
    """Rename slide parts to a contiguous ``slide1..slideN`` sequence.

    ``python-pptx`` derives the next slide part name from the number of
    slides, so a gap left by deleting a slide can cause a name collision on
    save. Renumbering keeps part names unique.
    """
    sld_ids = _sld_id_elements(prs)
    prs.part.rename_slide_parts([sld_id.rId for sld_id in sld_ids])


def _move_slide_id(prs: Presentation, from_index: int, to_index: int) -> None:
    """Move a ``<p:sldId>`` element within ``sldIdLst``."""
    sld_id_lst = prs.slides._sldIdLst
    sld_id = _sld_id_elements(prs)[from_index]
    sld_id_lst.remove(sld_id)

    remaining = _sld_id_elements(prs)
    if to_index < len(remaining):
        remaining[to_index].addprevious(sld_id)
    else:
        sld_id_lst.append(sld_id)


def _copy_relationships(source_part, dest_part) -> dict[str, str]:
    """Copy relationships (images, links, ...) and return an old->new rId map."""
    rid_map: dict[str, str] = {}
    for rId, rel in source_part.rels.items():
        if rel.reltype in _SKIP_RELTYPES:
            continue
        if rel.is_external:
            new_rId = dest_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rId = dest_part.relate_to(rel.target_part, rel.reltype)
        rid_map[rId] = new_rId
    return rid_map


def _remap_rels(element, rid_map: dict[str, str]) -> None:
    """Rewrite relationship IDs inside a copied XML subtree."""
    if not rid_map:
        return
    for el in element.iter():
        for attr in _REL_ATTRS:
            qn_attr = qn(attr)
            old = el.get(qn_attr)
            if old is not None and old in rid_map:
                el.set(qn_attr, rid_map[old])


def add_slide(
    prs: Presentation,
    layout_name: str | None = None,
    index: int | None = None,
) -> tuple[Slide, int]:
    """Add a slide and optionally insert it at ``index``.

    Returns the new slide and its final zero-based index.
    """
    count = len(prs.slides)
    if index is not None and (index < 0 or index > count):
        raise SlideIndexError(
            f"Insert index {index} is out of range (0..{count}); "
            f"the presentation has {count} slide(s)."
        )

    layout = resolve_layout(prs, layout_name)
    _normalize_slide_partnames(prs)
    slide = prs.slides.add_slide(layout)
    final_index = len(prs.slides) - 1

    if index is not None and index != final_index:
        _move_slide_id(prs, final_index, index)
        final_index = index

    return slide, final_index


def duplicate_slide(
    prs: Presentation,
    index: int,
    new_index: int | None = None,
) -> tuple[Slide, int]:
    """Duplicate a slide, copying its shapes, background and images.

    Returns the duplicated slide and its final zero-based index.
    """
    _validate_index(prs, index)
    count = len(prs.slides)
    if new_index is not None and (new_index < 0 or new_index > count):
        raise SlideIndexError(
            f"Destination index {new_index} is out of range (0..{count})."
        )

    source = prs.slides[index]
    _normalize_slide_partnames(prs)
    dest = prs.slides.add_slide(source.slide_layout)

    # Copy the source's relationships (images, hyperlinks, ...) so the copied
    # XML keeps pointing at valid targets in the new slide part.
    rid_map = _copy_relationships(source.part, dest.part)

    # Mutate the destination's existing cSld/spTree in place. Replacing the
    # whole cSld would leave the cached `Slide.shapes` proxy pointing at the
    # old (empty) shape tree.
    src_cSld = source._element.cSld
    dest_cSld = dest._element.cSld

    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is not None:
        new_bg = copy.deepcopy(src_bg)
        _remap_rels(new_bg, rid_map)
        dest_bg = dest_cSld.find(qn("p:bg"))
        if dest_bg is not None:
            dest_cSld.replace(dest_bg, new_bg)
        else:
            dest_cSld.insert(0, new_bg)

    dest_spTree = dest_cSld.spTree
    for child in list(dest_spTree):
        dest_spTree.remove(child)
    for child in src_cSld.spTree:
        new_el = copy.deepcopy(child)
        _remap_rels(new_el, rid_map)
        dest_spTree.append(new_el)

    final_index = len(prs.slides) - 1
    if new_index is not None and new_index != final_index:
        _move_slide_id(prs, final_index, new_index)
        final_index = new_index

    return dest, final_index


def delete_slide(prs: Presentation, index: int) -> None:
    """Delete the slide at ``index`` and drop its relationship."""
    _validate_index(prs, index)
    sld_id = _sld_id_elements(prs)[index]
    rId = sld_id.rId
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(sld_id)
    _normalize_slide_partnames(prs)


def move_slide(prs: Presentation, from_index: int, to_index: int) -> None:
    """Move the slide at ``from_index`` to ``to_index``."""
    _validate_index(prs, from_index)
    count = len(prs.slides)
    if to_index < 0 or to_index >= count:
        raise SlideIndexError(
            f"Destination index {to_index} is out of range (0..{count - 1}); "
            f"the presentation has {count} slide(s)."
        )
    if from_index == to_index:
        return
    _move_slide_id(prs, from_index, to_index)
