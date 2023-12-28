from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.draw import Draw, DrawDoc
from ooodev.utils.lo import Lo


def test_slides_draw(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc(loader))
    try:
        assert isinstance(doc, DrawDoc)
        assert len(doc.slides) == 1
        slide = doc.slides[0]
        assert slide is not None

        # master pages do not count as slides
        page = doc.insert_master_page(idx=0)
        assert page is not None
        assert len(doc.slides) == 1

        new_slide = doc.slides.insert_slide(idx=1)
        assert len(doc.slides) == 2
        new_slide.set_name("New Slide")

        slides2 = doc.get_slides()
        assert slides2 is doc.slides

        latest_slide = doc.insert_slide(idx=-1)
        assert len(doc.slides) == 3
        latest_slide.set_name("Latest Slide")

        last_slide = doc.slides[-1]
        assert last_slide.get_name() == "Latest Slide"

        second_last = doc.slides[-2]
        assert second_last.get_name() == "New Slide"

        doc.slides.delete_slide(idx=-1)
        assert len(doc.slides) == 2
        last_slide = doc.slides[-1]
        assert last_slide.get_name() == "New Slide"
    finally:
        doc.close_doc()
