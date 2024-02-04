"""
Write 10 forumlae into a new Text doc and save as a pdf file.
"""

import pytest
import random
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])
import uno

from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI

from com.sun.star.text import XTextDocument


def test_shuffle_words(loader, copy_fix_writer, tmp_path_fn):
    visible = False
    delay = 0  # 4_000
    loop_delay = 100
    test_doc = copy_fix_writer("storyStart.doc")
    fnm_save = Path(tmp_path_fn, "storyStart.doc")
    Write.open_doc(fnm=test_doc, loader=loader)
    # could have captured doc as: doc = Write.open_doc(fnm=test_doc, loader=loader)
    # just confirming that Lo.XSCRIPTCONTEXT is working.
    # Document is not visible at this point os it is not available via Lo.XSCRIPTCONTEXT.getDocument()
    # or Lo.this_component
    doc = Lo.lo_component
    assert doc is not None

    try:
        if visible:
            GUI.set_visible(visible, doc)
        apply_shuffle(doc, loop_delay, visible)

        Lo.delay(delay)
        Write.save_doc(text_doc=doc, fnm=fnm_save)
        assert fnm_save.exists()
    finally:
        Lo.close_doc(doc, False)


def apply_shuffle(doc: XTextDocument, delay: int, visible: bool) -> None:
    doc_text = doc.getText()
    if visible:
        cursor = Write.get_view_cursor(doc)
    else:
        cursor = Write.get_cursor(doc)

    word_cursor = Write.get_word_cursor(doc)
    word_cursor.gotoStart(False)  # go to start of text

    while True:
        word_cursor.gotoNextWord(True)

        # move the text view cursor, and highlight the current word
        cursor.gotoRange(word_cursor.getStart(), False)
        cursor.gotoRange(word_cursor.getEnd(), True)
        curr_word = word_cursor.getString()

        # get whitespace padding amounts
        c_len = len(curr_word)
        curr_word = curr_word.lstrip()
        l_pad = c_len - len(curr_word)  # left whitespace padding amount
        curr_word = curr_word.rstrip()
        r_pad = c_len - len(curr_word) - l_pad  # right whitespace padding ammount
        if len(curr_word) > 0:
            pad_l = " " * l_pad  # recreate left padding
            pad_r = " " * r_pad  # recreate right padding
            Lo.delay(delay)
            mid_shuff = mid_shuffle(curr_word)
            doc_text.insertString(word_cursor, f"{pad_l}{mid_shuff}{pad_r}", True)

        if word_cursor.gotoNextWord(False) is False:
            break
    word_cursor.gotoStart(False)  # go to start of text
    cursor.gotoStart(False)


def mid_shuffle(s: str) -> str:
    s_len = len(s)
    if s_len <= 3:  # not long enough
        return s

    # extract middle of the string for shuffling
    mid = s[1 : s_len - 1]

    # shuffle a list of characters  made from the middle
    mid_lst = list(mid)
    random.shuffle(mid_lst)

    # rebuild string, adding back the first and last letters
    # punctuation may be first or last char
    return f"{s[:1]}{''.join(mid_lst)}{s[-1:]}"
