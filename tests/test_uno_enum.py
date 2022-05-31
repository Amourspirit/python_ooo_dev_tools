from __future__ import annotations
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.uno_util import UnoEnum



def test_uno_enum_singleton() -> None:
    ue = UnoEnum("com.sun.star.sheet.FillMode")
    LINEAR = ue.LINEAR
    assert LINEAR.value == "LINEAR"
    LINEAR = ue.LINEAR
    assert LINEAR.value == "LINEAR"
    with pytest.raises(AttributeError):
        v = ue.Not_Existing
    
    FillMode = UnoEnum("com.sun.star.sheet.FillMode")
    assert FillMode is ue
    # checking for a valid attribue acutally adds it.
    assert hasattr(ue, "AUTO")
    assert FillMode.DATE.value == "DATE"
    assert hasattr(ue, "DATE")
    assert ue.DATE.value == "DATE"
    assert hasattr(ue, "NOT_EXISTING") == False

def test_import_same() -> None:
    from com.sun.star.sheet.FillDirection import TO_RIGHT, TO_LEFT, TO_TOP
    fd = UnoEnum("com.sun.star.sheet.FillDirection")
    assert fd.TO_RIGHT == TO_RIGHT
    assert fd.TO_LEFT == TO_LEFT
    assert fd.TO_TOP == TO_TOP
    