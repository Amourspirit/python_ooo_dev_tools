"""
Write 10 forumlae into a new Text doc and save as a pdf file.
"""
from typing import Sequence
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])
import uno

from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI


from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange
from com.sun.star.util import XReplaceable
from com.sun.star.util import XReplaceDescriptor
from com.sun.star.util import XSearchable
from com.sun.star.beans import XPropertySet


def test_text_replace(loader, copy_fix_writer, tmp_path_fn):
    visible = False
    delay = 0  # 2_000
    test_doc = copy_fix_writer("bigStory.doc")
    fnm_save = Path(tmp_path_fn, "story_replace.odt")
    doc = Write.open_doc(fnm=test_doc, loader=loader)

    try:
        if visible:
            GUI.set_visible(visible, doc)

        uk_words = ("colour", "neighbour", "centre", "behaviour", "metre", "through")
        us_words = ("color", "neighbor", "center", "behavior", "meter", "thru")

        if visible:
            words = (
                "(G|g)rit",
                "colou?r",
            )
            find_words(doc, words)

        num = replace_words(doc, uk_words, us_words)
        assert num == 4

        Lo.delay(delay)
        Write.save_doc(text_doc=doc, fnm=fnm_save)
        assert fnm_save.exists()
    finally:
        Lo.close_doc(doc, False)


def find_words(doc: XTextDocument, words: Sequence[str]) -> None:
    # get the view cursor and link the page cursor to it
    tvc = Write.get_view_cursor(doc)
    tvc.gotoStart(False)
    page_cursor = Write.get_page_cursor(tvc)
    searchable = Lo.qi(XSearchable, doc)
    srch_desc = searchable.createSearchDescriptor()

    for word in words:
        print(f"Searching for fist occurrence of '{word}'")
        srch_desc.setSearchString(word)

        srch_props = Lo.qi(XPropertySet, srch_desc, raise_err=True)
        srch_props.setPropertyValue("SearchRegularExpression", True)

        srch = searchable.findFirst(srch_desc)

        if srch is not None:
            match_tr = Lo.qi(XTextRange, srch)

            tvc.gotoRange(match_tr, False)
            print(f"  - found '{match_tr.getString()}'")
            print(f"    - on page {page_cursor.getPage()}")
            # tvc.gotoStart(True)
            tvc.goRight(len(match_tr.getString()), True)
            print(f"    - at char postion: {len(tvc.getString())}")
            Lo.delay(500)


def replace_words(doc: XTextDocument, old_words: Sequence[str], new_words: Sequence[str]) -> int:
    replaceable = Lo.qi(XReplaceable, doc, raise_err=True)
    replace_desc = Lo.qi(XReplaceDescriptor, replaceable.createSearchDescriptor())

    for old, new in zip(old_words, new_words):
        replace_desc.setSearchString(old)
        replace_desc.setReplaceString(new)
    return replaceable.replaceAll(replace_desc)
