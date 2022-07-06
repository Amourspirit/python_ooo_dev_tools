from unicodedata import name
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])
import uno

from com.sun.star.text import XTextDocument
from com.sun.star.util import XSearchable
from com.sun.star.text import XTextRange

from ooo.dyn.awt.font_slant import FontSlant  # enum

from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.props import Props
from ooodev.utils.color import CommonColor, Color


def test_italic_styler(loader, copy_fix_writer):
    visible = False
    delay = 0 #, 1_000
    test_doc = copy_fix_writer("cicero_dummy.odt")
    doc = Write.open_doc(test_doc, loader)
    try:
        if visible:
            GUI.set_visible(visible, doc)
        # lock screen updates for first test
        with Lo.ControllerLock():
            result = italicize_all(doc, "pleasure", CommonColor.RED)
            assert result == 17
        Lo.delay(delay)
        result = italicize_all(doc, "pain", CommonColor.GREEN)
        assert result == 15
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)


def italicize_all(doc: XTextDocument, phrase: str, color:Color ) -> int:
    # cursor = Write.get_view_cursor(doc) # can be used when visible
    cursor = Write.get_cursor(doc)
    cursor.gotoStart(False)
    page_cursor = Write.get_page_cursor(doc)
    result = 0
    try:
        xsearchable = Lo.qi(XSearchable, doc, True)
        srch_desc = xsearchable.createSearchDescriptor()
        print(f"Searching for all occurrences of '{phrase}'")
        pharse_len = len(phrase)
        srch_desc.setSearchString(phrase)
        # for props see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html
        Props.set_property(obj=srch_desc, name="SearchCaseSensitive", value=False)
        Props.set_property(obj=srch_desc, name="SearchWords", value=True) # If TRUE, only complete words will be found.    

        matches = xsearchable.findAll(srch_desc)
        result = matches.getCount()

        print(f"No. of matches: {result}")

        for i in range(result):
            match_tr = Lo.qi(XTextRange, matches.getByIndex(i))
            if match_tr is not None:
                cursor.gotoRange(match_tr, False)
                print(f"  - found: '{match_tr.getString()}'")
                print(f"    - on page {page_cursor.getPage()}")
                cursor.gotoStart(True)
                print(f"    - starting at char position: {len(cursor.getString()) - pharse_len}")

                Props.set_properties(
                    obj=match_tr, names=("CharColor", "CharPosture"), vals=(color, FontSlant.ITALIC)
                )

    except Exception as e:
        raise
    return result
