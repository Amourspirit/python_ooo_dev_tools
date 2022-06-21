from typing import TYPE_CHECKING, cast
import pytest
from pathlib import Path
# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.utils.info import Info
from ooodev.office.write import Write
from ooodev.utils.gui import GUI

from com.sun.star.text import XTextDocument
from com.sun.star.text import XText
from com.sun.star.text import XTextCursor
from com.sun.star.text import XTextRange


if TYPE_CHECKING:
    from com.sun.star.text import TextCursor # service
    from com.sun.star.text import Paragraph # service
    from com.sun.star.text import TextPortion # servcie

def test_writer_scandal_start(loader, copy_fix_writer):
    visible = True
    delay = 2000

    # text file opens with each new line being considered a paragraph break.
    test_doc = copy_fix_writer("scandalStart.odt")
    doc = Write.open_doc(test_doc, loader)
    try:
        if visible:
            GUI.set_visible(visible, doc)
        # last_para = get_last_para(doc)
        # assert last_para == "At three o'clock precisely I was at Baker Street, but Holmes had not yet returned. The landlady informed me that he had left the house shortly after eight o'clock in the morning. I sat down beside the fire, however, with the intention of awaiting him, however long he might be."
       
        # para_count = count_para(doc)
        # paracount, sentencecount = count_paras_and_sentences(doc)
        # assert para_count == 7

        enumerate_text_sections(doc)

        # sent_count = count_sentences(doc)

        # word_count = count_words(doc)
        # assert word_count == 152
        
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)

def get_last_para(doc: XTextDocument) -> str:
    # tvc = Write.get_view_cursor(doc)
    para_cursor = Write.get_paragraph_cursor(doc)
    para_cursor.gotoStart(False)
    
    curr_para = ''
    while True:
        para_cursor.gotoEndOfParagraph(True) # select all of paragraph
        curr_para = para_cursor.getString()
        if para_cursor.gotoNextParagraph(False) is False:
            break
    return curr_para

def count_words(doc: XTextDocument) -> int:
    word_cursor = Write.get_word_cursor(doc)
    word_cursor.gotoStart(False)
    
    count = 0
    while True:
        word_cursor.gotoEndOfWord(True)
        curr_word = word_cursor.getString()
        if len(curr_word) > 0:
            count += 1
        if word_cursor.gotoNextWord(False) is False:
            break
    return count

def count_sentences(doc: XTextDocument) -> int:
    sent_cursor = Write.get_sentence_cursor(doc)
    sent_cursor.gotoStart(False)
    
    count = 0
    loop_count = 0
    while True:
        loop_count += 1
        sent_cursor.gotoEndOfSentence(True)
        s = sent_cursor.getString()
        if len(s) > 0:
            count += 1
        if sent_cursor.gotoNextSentence(False) is False:
            break
        assert loop_count < 50
    return count

def count_para(doc: XTextDocument) -> int:
    # see: OOME_4_0.pdf pg: 393
    # see: https://flylib.com/books/en/4.290.1.130/1/
    para_cursor = Write.get_paragraph_cursor(doc)
    para_cursor.gotoStart(False)
    p_count = 0
    loop_count = 0
    while True:
        loop_count += 1
        # para_cursor.gotoStartOfParagraph(False)
        # The cursor is already positioned at the start
        # of the current paragraph, so select the entire paragraph.
        para_cursor.gotoEndOfParagraph(True)
        para_str = para_cursor.getString()
        if len(para_str) > 0:
            p_count += 1
        if para_cursor.gotoNextParagraph(False) is False:
            break
        assert loop_count < 50
    return p_count


def count_Sentences(cursor: "TextCursor", count=0) -> int:
    while not cursor.isEndOfParagraph():
        count += 1
        cursor.gotoNextSentence( False )
        cursor.gotoEndOfSentence( False )
    return count


def count_paras_and_sentences(doc: XTextDocument):
    # https://ask.libreoffice.org/t/cursor-gotonextsentence-failing/23129/3
    p_cursor = Write.get_paragraph_cursor(doc)
    assert Info.support_service(p_cursor, "com.sun.star.text.TextCursor")
    cursor = cast("TextCursor", p_cursor)
    cursor.gotoStart(False)
    sentencecount = 0
    paracount = 0
    while cursor.gotoNextParagraph(False):
        sentencecount += count_Sentences(cursor, sentencecount)
        paracount += 1

    return paracount, sentencecount

def count_cursor_words(cursor:XTextCursor) -> int:
    word_cursor = Write.get_word_cursor(cursor)
    word_cursor.gotoStart(False)
    
    count = 0
    while True:
        word_cursor.gotoEndOfWord(True)
        curr_word = word_cursor.getString()
        if len(curr_word) > 0:
            count += 1
        if word_cursor.gotoNextWord(False) is False:
            break
    return count

def enumerate_text_sections(doc: XTextDocument):
    e = Write.get_enumeration(doc)
    p_count = 0
    w_count = 0
    xtext = doc.getText()
    while e.hasMoreElements():
        para = e.nextElement()
        if Info.support_service(para, "com.sun.star.text.Paragraph"):
            p_enum = cast("Paragraph", para)
            p_range = Lo.qi(XTextRange, p_enum)
            r_str = p_range.getString()
            
            if r_str == "":
                continue
            p_count += 1
            s = f"{p_count}:"
            
            cursor = Write.get_cursor(rng=p_range, txt=xtext)
            w_count += count_cursor_words(cursor=cursor)
            
            t_e = p_enum.createEnumeration()
            while t_e.hasMoreElements():
                # elemet here are portions. For instance of a word is highlighted then it will be a seperate portion
                para_section = cast("TextPortion", t_e.nextElement())
                
                
                ps = para_section.getString()
                # s = s & oParSection.TextPortionType & ":"
                s = f"{s}{para_section.TextPortionType}:"
        print(s)