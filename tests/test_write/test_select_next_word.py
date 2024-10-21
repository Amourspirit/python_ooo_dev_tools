import pytest
import sys

if __name__ == "__main__":
    pytest.main([__file__])

from com.sun.star.document import MacroExecMode

from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.gui.gui import GUI

LAST_PARA = """
At three o'clock precisely I was at Baker Street, but Holmes had not yet returned. The landlady informed me that he had left the house shortly after eight o'clock in the morning. I sat down beside the fire, however, with the intention of awaiting him, however long he might be.
"""

# on windows getting Fatal Python error: Aborted even though the test runs fine when run by itself.


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows in a group")
def test_select_next_word(loader, copy_fix_writer):
    from ooodev.utils.props import Props

    # text file opens with each new line being considered a paragraph break.

    # test require Writer be visible
    visible = True
    delay = 300
    test_doc = copy_fix_writer("scandalStart.odt")
    props = Props.make_props(Hidden=True, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN)
    doc = Write.open_doc(fnm=test_doc, loader=loader, props=props)
    try:
        if visible:
            GUI.set_visible(visible, doc)

        # make sure when doc opens cursor is set to start
        vc = Write.get_view_cursor(doc)
        vc.gotoStart(False)
        vc.collapseToStart()

        assert Write.is_anything_selected(doc) is False
        Write.select_next_word(doc)
        assert Write.is_anything_selected(doc)

        s = Write.get_selected_text_str(doc)
        assert s == "Title: "
        Lo.delay(delay)

        Write.select_next_word(doc)
        s = Write.get_selected_text_str(doc)
        assert s == "A "
        Lo.delay(delay)

        # no deselect
        Write.select_next_word(doc)
        s = Write.get_selected_text_str(doc)
        assert s == "Scandal "
        Lo.delay(delay)

        # Somewhere after LibreOffice 7.6 select_next_word went from selecting the word with the punctuation to the word and the punctuation being a next selection.
        # This is a change in behavior.
        # So, the next selection will be "returned" instead of "returned." and the next selection will be "." instead of "The ".
        # For this reason the next part is no longer tested.

        # jump to last paragraph
        # vc.goRight(554, False)
        # words = LAST_PARA.split()
        # for i in range(50):  # 50 Max
        #     Write.select_next_word(doc)
        #     s = Write.get_selected_text_str(doc).strip()
        #     assert s == words[i]

        #     Lo.delay(100)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)
