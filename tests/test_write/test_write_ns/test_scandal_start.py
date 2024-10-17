from typing import TYPE_CHECKING
import pytest
import sys

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.write import Write
from ooodev.write import WriteDoc
from ooodev.utils.selection import WordTypeEnum


if TYPE_CHECKING:
    pass  # service

# Other resources:  https://flylib.com/books/en/4.290.1.130/1/
#                       OOME_4_0.pdf pg: 393
#                       Everything Cursors
#                   https://ask.libreoffice.org/t/cursor-gotonextsentence-failing/23129/3
#                       Demonstrates a pythonic way of enumerating paragraphs and sentences.
#                       I found count_Sentences to not work, got something like 1500 sentences on scandalStart.odt, Way too high

# on windows getting Fatal Python error: Aborted even though the test runs fine when run by itself.


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows in a group")
def test_writer_scandal_start(loader, copy_fix_writer):
    visible = True
    delay = 2000

    # text file opens with each new line being considered a paragraph break.
    test_doc = copy_fix_writer("scandalStart.odt")
    doc = WriteDoc(Write.open_doc(test_doc, loader))
    try:
        if visible:
            doc.set_visible()

        para_count, word_count, parts_count = enumerate_text_sections(doc)
        assert para_count == 7
        assert word_count == 153
        assert parts_count == 11
        # get_last_para(doc)

        Lo.delay(delay)
    finally:
        doc.close_doc()


def get_last_para(doc: WriteDoc) -> str:
    # tvc = Write.get_view_cursor(doc)
    para_cursor = doc.get_paragraph_cursor()
    para_cursor.goto_start()

    curr_para = ""
    while True:
        para_cursor.goto_end_of_paragraph(True)
        curr_para = para_cursor.get_string()
        if para_cursor.goto_next_paragraph(False) is False:
            break
    return curr_para


def enumerate_text_sections(doc: WriteDoc):
    doc_text = doc.get_text()
    p_count = 0
    w_count = 0
    parts_count = 0
    paragraphs = doc_text.get_paragraphs()
    # paragraphs = doc.get_text_paragraphs()
    s = ""
    for para in paragraphs:
        r_str = para.get_string()

        if r_str == "":
            continue
        p_count += 1
        s = f"{p_count}:"

        # cursor = Write.get_cursor(rng=p_range, txt=xtext)
        # w_count += count_cursor_words(cursor=cursor)
        w_count += Write.get_word_count_ooo(text=r_str, word_type=WordTypeEnum.WORD_COUNT)

        portions = para.get_text_portions()
        for portion in portions:
            parts_count += 1
            s = f"{s}{portion.text_portion_type}:"
        print(s)
    return p_count, w_count, parts_count
