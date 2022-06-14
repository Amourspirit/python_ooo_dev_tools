import pytest
from pathlib import Path
# from ooodev.office.write import Write
import uno

if __name__ == "__main__":
    pytest.main([__file__])

def test_ControlCharacter():
    from ooodev.office.write import Write
    from ooodev.utils.uno_const import UnoConst
    from com.sun.star.text import ControlCharacter
    assert Write.ControlCharacter.APPEND_PARAGRAPH == ControlCharacter.APPEND_PARAGRAPH
    assert Write.ControlCharacter.HARD_HYPHEN == ControlCharacter.HARD_HYPHEN
    assert Write.ControlCharacter.HARD_SPACE == ControlCharacter.HARD_SPACE
    assert Write.ControlCharacter.HARD_HYPHEN == ControlCharacter.HARD_HYPHEN
    assert Write.ControlCharacter.LINE_BREAK == ControlCharacter.LINE_BREAK
    assert Write.ControlCharacter.PARAGRAPH_BREAK == ControlCharacter.PARAGRAPH_BREAK
    assert Write.ControlCharacter.SOFT_HYPHEN == ControlCharacter.SOFT_HYPHEN
    
    assert Write.ControlCharacter.SOFT_HYPHEN == ControlCharacter.SOFT_HYPHEN
    
    cc = UnoConst("com.sun.star.text.ControlCharacter")
    assert cc is Write.ControlCharacter
    assert cc.LINE_BREAK == 1
    