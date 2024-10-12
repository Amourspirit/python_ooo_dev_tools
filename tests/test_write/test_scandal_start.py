from typing import TYPE_CHECKING, cast
import sys
import pytest
from pathlib import Path

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

import uno
from com.sun.star.document import MacroExecMode

from ooodev.loader.lo import Lo
from ooodev.utils.info import Info
from ooodev.office.write import Write
from ooodev.gui.gui import GUI
from ooodev.utils.selection import WordTypeEnum

from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange


if TYPE_CHECKING:
    from com.sun.star.text import Paragraph  # service
    from com.sun.star.text import TextPortion  # service

# Other resources:  https://flylib.com/books/en/4.290.1.130/1/
#                       OOME_4_0.pdf pg: 393
#                       Everything Cursors
#                   https://ask.libreoffice.org/t/cursor-gotonextsentence-failing/23129/3
#                       Demonstrates a pythonic way of enumerating paragraphs and sentences.
#                       I found count_Sentences to not work, got something like 1500 sentences on scandalStart.odt, Way too high

# on windows getting Fatal Python error: Aborted even though the test runs fine when run by itself.


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows in a group")
def test_writer_scandal_start(loader, copy_fix_writer):
    from ooodev.utils.props import Props

    visible = True
    delay = 2000

    # text file opens with each new line being considered a paragraph break.
    test_doc = copy_fix_writer("scandalStart.odt")
    props = Props.make_props(Hidden=True, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN)
    doc = Write.open_doc(fnm=test_doc, loader=loader, props=props)
    try:
        if visible:
            GUI.set_visible(visible, doc)

        para_count, word_count, parts_count = enumerate_text_sections(doc)
        assert para_count == 7
        assert word_count == 153
        assert parts_count == 11

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)


def get_last_para(doc: XTextDocument) -> str:
    # tvc = Write.get_view_cursor(doc)
    para_cursor = Write.get_paragraph_cursor(doc)
    para_cursor.gotoStart(False)

    curr_para = ""
    while True:
        para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
        curr_para = para_cursor.getString()
        if para_cursor.gotoNextParagraph(False) is False:
            break
    return curr_para


def enumerate_text_sections(doc: XTextDocument):
    e = Write.get_enumeration(doc)
    p_count = 0
    w_count = 0
    parts_count = 0
    # xtext = doc.getText()
    s = ""
    while e.hasMoreElements():
        para = e.nextElement()
        if Info.support_service(para, "com.sun.star.text.Paragraph"):
            p_enum = cast("Paragraph", para)
            p_range = Lo.qi(XTextRange, p_enum, True)
            r_str = p_range.getString()

            if r_str == "":
                continue
            p_count += 1
            s = f"{p_count}:"

            # cursor = Write.get_cursor(rng=p_range, txt=xtext)
            # w_count += count_cursor_words(cursor=cursor)
            w_count += Write.get_word_count_ooo(text=r_str, word_type=WordTypeEnum.WORD_COUNT)

            t_e = p_enum.createEnumeration()
            while t_e.hasMoreElements():
                # element here are portions. For instance of a word is highlighted then it will be a separate portion
                parts_count += 1
                para_section = cast("TextPortion", t_e.nextElement())
                ps = para_section.getString()
                # s = s & oParSection.TextPortionType & ":"
                s = f"{s}{para_section.TextPortionType}:"
        # print(s)
    return p_count, w_count, parts_count
