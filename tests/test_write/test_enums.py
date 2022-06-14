import pytest
from pathlib import Path
# from ooodev.office.write import Write
import uno

if __name__ == "__main__":
    pytest.main([__file__])

def test_BreakType():
    from ooodev.office.write import Write
    from com.sun.star.style.BreakType import (
        NONE,
        COLUMN_BEFORE,
        COLUMN_AFTER,
        COLUMN_BOTH,
        PAGE_BEFORE,
        PAGE_AFTER,
        PAGE_BOTH
    )
    assert Write.BreakType.NONE == NONE
    assert Write.BreakType.COLUMN_BEFORE == COLUMN_BEFORE
    assert Write.BreakType.COLUMN_AFTER == COLUMN_AFTER
    assert Write.BreakType.COLUMN_BOTH == COLUMN_BOTH
    assert Write.BreakType.PAGE_BEFORE == PAGE_BEFORE
    assert Write.BreakType.PAGE_AFTER == PAGE_AFTER
    assert Write.BreakType.PAGE_BOTH == PAGE_BOTH

def test_ParagraphAdjust():
    from ooodev.office.write import Write
    from com.sun.star.style.ParagraphAdjust import (
        LEFT,
        RIGHT,
        BLOCK,
        CENTER,
        STRETCH
    )
    assert Write.ParagraphAdjust.LEFT == LEFT
    assert Write.ParagraphAdjust.RIGHT == RIGHT
    assert Write.ParagraphAdjust.BLOCK == BLOCK
    assert Write.ParagraphAdjust.CENTER == CENTER
    assert Write.ParagraphAdjust.STRETCH == STRETCH

def test_FontSlant():
    from ooodev.office.write import Write
    from com.sun.star.awt.FontSlant import (
        NONE,
        OBLIQUE,
        ITALIC,
        DONTKNOW,
        REVERSE_OBLIQUE,
        REVERSE_ITALIC
    )
    assert Write.FontSlant.NONE == NONE
    assert Write.FontSlant.OBLIQUE == OBLIQUE
    assert Write.FontSlant.ITALIC == ITALIC
    assert Write.FontSlant.DONTKNOW == DONTKNOW
    assert Write.FontSlant.REVERSE_OBLIQUE == REVERSE_OBLIQUE
    assert Write.FontSlant.REVERSE_ITALIC == REVERSE_ITALIC

def test_PageNumberType():
    from ooodev.office.write import Write
    from com.sun.star.text.PageNumberType import (
        PREV,
        CURRENT,
        NEXT
    )
    assert Write.PageNumberType.PREV == PREV
    assert Write.PageNumberType.CURRENT == CURRENT
    assert Write.PageNumberType.NEXT == NEXT

def test_DictionaryType():
    from ooodev.office.write import Write
    from com.sun.star.linguistic2.DictionaryType import (
        POSITIVE,
        NEGATIVE,
        MIXED
    )
    assert Write.DictionaryType.POSITIVE == POSITIVE
    assert Write.DictionaryType.NEGATIVE == NEGATIVE
    assert Write.DictionaryType.MIXED == MIXED # Deprecated:

def test_PaperFormat():
    from ooodev.office.write import Write
    from com.sun.star.view.PaperFormat import (
        A3,
        A4,
        A5,
        B4,
        B5,
        LETTER,
        LEGAL,
        TABLOID,
        USER
    )
    assert Write.PaperFormat.A3 == A3
    assert Write.PaperFormat.A4 == A4
    assert Write.PaperFormat.A5 == A5
    assert Write.PaperFormat.B4 == B4
    assert Write.PaperFormat.B5 == B5
    assert Write.PaperFormat.LETTER == LETTER
    assert Write.PaperFormat.LEGAL == LEGAL
    assert Write.PaperFormat.TABLOID == TABLOID
    assert Write.PaperFormat.USER == USER

def test_TextContentAnchorType():
    from ooodev.office.write import Write
    from com.sun.star.text.TextContentAnchorType import (
        AT_PARAGRAPH,
        AS_CHARACTER,
        AT_PAGE,
        AT_FRAME,
        AT_CHARACTER
    )
    assert Write.TextContentAnchorType.AT_PARAGRAPH == AT_PARAGRAPH
    assert Write.TextContentAnchorType.AS_CHARACTER == AS_CHARACTER
    assert Write.TextContentAnchorType.AT_PAGE == AT_PAGE
    assert Write.TextContentAnchorType.AT_FRAME == AT_FRAME
    assert Write.TextContentAnchorType.AT_CHARACTER == AT_CHARACTER