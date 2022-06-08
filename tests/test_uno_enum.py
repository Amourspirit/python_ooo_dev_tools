from __future__ import annotations
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.uno_enum import UnoEnum
from typing import TYPE_CHECKING, cast

# uno does not have enum import but rather enum values.
# that means:
#   from com.sun.star.sheet import FillMode # import error
#   from com.sun.star.sheet.FillMode import LINEAR, DATE # this import is fine
# by using TYPE_CHECKING, cast and wrapping import name in string we have the same
# experience as if we had imported an enum
# Example:
# if TYPE_CHECKING:
#        from com.sun.star.sheet import FillMode as UnoFillMode
# FillMode = cast("UnoFillMode", UnoEnum("com.sun.star.sheet.FillMode")) # works like a regular enum for typings

def test_uno_enum_singleton() -> None:
    if TYPE_CHECKING:
        from com.sun.star.sheet import FillMode as UnoFillMode
    ue = cast("UnoFillMode", UnoEnum("com.sun.star.sheet.FillMode"))
    LINEAR = ue.LINEAR
    assert LINEAR.value == "LINEAR"
    LINEAR = ue.LINEAR
    assert LINEAR.value == "LINEAR"
    with pytest.raises(AttributeError):
        v = ue.Not_Existing
    
    FillMode = cast("UnoFillMode", UnoEnum("com.sun.star.sheet.FillMode"))
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
    