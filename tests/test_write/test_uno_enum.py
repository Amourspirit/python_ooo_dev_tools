import pytest
from pathlib import Path
# from ooodev.office.write import Write
import uno

if __name__ == "__main__":
    pytest.main([__file__])

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
