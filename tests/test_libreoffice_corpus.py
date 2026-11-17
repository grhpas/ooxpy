"""oox's readers against files LibreOffice wrote, with known contents.

The files and ``truth.json`` come from ``python -m tools.locorpus``: flat
OpenDocument sources written there, converted to .docx by LibreOffice. The truth
is what the sources contain, plus LibreOffice's own reading of each .docx as
text (its list labels).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from oox import open_document, teletype
from oox.walk import iter_story

DATA = Path(__file__).parent / "data" / "libreoffice"
TRUTH = json.loads((DATA / "truth.json").read_text(encoding="utf-8"))["documents"]


def doc(name: str):
    return open_document(DATA / name)


def test_list_labels_match_the_source_and_libreoffice():
    truth = TRUTH["lists.docx"]
    labels = [item.label for item in doc("lists.docx").list_labels()]
    assert labels == truth["labels"]
    assert labels == [line.split()[0] for line in truth["libreoffice_text"]]


def test_list_items_say_where_they_come_from():
    items = doc("lists.docx").list_labels()
    headings = [i for i in items if i.label in ("1", "1.1", "1.2", "2")]
    assert [i.level for i in headings] == [0, 1, 1, 0]
    assert all(i.is_bullet for i in items if i.label == "•")
    assert {i.story for i in items} == {"body"}


def test_sections_and_page_orientation():
    truth = TRUTH["sections.docx"]
    sections = doc("sections.docx").sections
    assert len(sections) == truth["sections"]
    assert [s.orientation for s in sections] == truth["orientations"]


def test_pictures_text_boxes_and_tables():
    truth = TRUTH["frames.docx"]
    d = doc("frames.docx")
    pictures = d.pictures()
    assert len(pictures) == truth["pictures"]
    assert all(d.image_of(p).content_type == "image/png" for p in pictures)
    assert len(d.text_boxes()) == truth["text_boxes"]
    assert len(d.tables()) == truth["tables"]
    text = teletype.extractText(d.body)
    for once in truth["once"]:
        assert text.count(once) == 1


@pytest.mark.parametrize("name", sorted(TRUTH))
def test_reading_resolves_every_paragraph(name):
    d = doc(name)
    fmt = d.formatting()
    for p in iter_story(d.body, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
        assert fmt.paragraph(p).line_spacing.rule in ("auto", "exact", "at_least")
        assert fmt.mark(p).size > 0
