import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from com.sun.star.util import XSearchable
from com.sun.star.text import XTextRange

from ooo.dyn.awt.font_slant import FontSlant  # enum

from ooodev.loader.lo import Lo
from ooodev.write import Write
from ooodev.write import WriteDoc
from ooodev.utils.props import Props
from ooodev.utils.color import CommonColor, Color


def test_italic_styler(loader, copy_fix_writer):
    visible = False
    delay = 0  # , 1_000
    test_doc = copy_fix_writer("cicero_dummy.odt")
    doc = WriteDoc(Write.open_doc(test_doc, loader))
    try:
        if visible:
            doc.set_visible()
        # lock screen updates for first test
        with Lo.ControllerLock():
            result = italicize_all(doc, "pleasure", CommonColor.RED)
            assert result == 17
        Lo.delay(delay)
        result = italicize_all(doc, "pain", CommonColor.GREEN)
        assert result == 15
        Lo.delay(delay)
    finally:
        doc.close_doc()


def italicize_all(doc: WriteDoc, phrase: str, color: Color) -> int:
    # cursor = Write.get_view_cursor(doc) # can be used when visible
    cursor = doc.get_cursor()
    cursor.goto_start()
    page_cursor = doc.get_view_cursor()
    result = 0
    try:
        searchable = doc.qi(XSearchable, True)
        srch_desc = searchable.createSearchDescriptor()
        print(f"Searching for all occurrences of '{phrase}'")
        phrase_len = len(phrase)
        srch_desc.setSearchString(phrase)
        # for props see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html
        Props.set(srch_desc, SearchCaseSensitive=False)
        Props.set(srch_desc, SearchWords=True)  # If TRUE, only complete words will be found.

        matches = searchable.findAll(srch_desc)
        result = matches.getCount()

        print(f"No. of matches: {result}")

        for i in range(result):
            match_tr = Lo.qi(XTextRange, matches.getByIndex(i))
            if match_tr is not None:
                cursor.goto_range(match_tr, False)
                print(f"  - found: '{match_tr.getString()}'")
                print(f"    - on page {page_cursor.get_page()}")
                cursor.goto_start(True)
                print(f"    - starting at char position: {len(cursor.get_string()) - phrase_len}")

                Props.set_properties(obj=match_tr, names=("CharColor", "CharPosture"), vals=(color, FontSlant.ITALIC))

    except Exception:
        raise
    return result
