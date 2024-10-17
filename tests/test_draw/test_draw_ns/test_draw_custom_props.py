from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.draw import DrawDoc
from ooodev.utils.helper.dot_dict import DotDict


def test_draw_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:
        pth = Path(tmp_path, "test_draw_custom_props.odg")
        doc = DrawDoc.create_doc(loader)

        dd = DotDict()
        dd.doc_name = "Who knows"
        dd.doc_num = 1
        dd.boolean = True
        dd.nothing = ""
        page = doc.slides[0]

        page.set_custom_properties(dd)

        dd = page.get_custom_properties()
        assert dd.doc_name == "Who knows"
        assert dd.doc_num == 1
        assert dd.boolean is True
        assert dd.nothing == ""

        page.draw_rectangle(0, 0, 100, 100)

        page.set_custom_property("custom", "custom_value")

        page2 = doc.slides.insert_slide(1)
        dd = DotDict()
        dd.doc_name = "Who knew"
        dd.doc_num = 2
        dd.boolean = True
        page2.set_custom_properties(dd)

        doc.save_doc(pth)
        doc.close()

        doc = DrawDoc.open_doc(pth)
        page = doc.slides[0]
        dd = page.get_custom_properties()
        assert dd.doc_name == "Who knows"
        assert dd.doc_num == 1
        assert dd.boolean is True
        assert dd.custom == "custom_value"
        assert len(dd) == 5

        page2 = doc.slides[1]
        dd = page2.get_custom_properties()
        assert dd.doc_name == "Who knew"
        assert dd.doc_num == 2
        assert dd.boolean is True
        assert len(dd) == 3

    finally:
        if doc:
            doc.close()


def test_doc_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:
        pth = Path(tmp_path, "test_doc_custom_props.odg")
        doc = DrawDoc.create_doc(loader)
        dd = DotDict()
        dd.test1 = 10
        dd.test2 = "Hello World"
        dd.test3 = None
        doc.set_custom_properties(dd)
        test1 = doc.get_custom_property("test1")
        assert test1 == 10
        test2 = doc.get_custom_property("test2")
        assert test2 == "Hello World"
        test3 = doc.get_custom_property("test3")
        assert test3 is None

        doc.save_doc(pth)
        doc.close()

        doc = DrawDoc.open_doc(pth)
        dd = doc.get_custom_properties()
        assert dd.test1 == 10
        assert dd.test2 == "Hello World"
        assert dd.test3 is None

        test1 = doc.get_custom_property("test1")
        assert test1 == 10
        test2 = doc.get_custom_property("test2")
        assert test2 == "Hello World"
        test3 = doc.get_custom_property("test3")
        assert test3 is None

    finally:
        if doc:
            doc.close()
