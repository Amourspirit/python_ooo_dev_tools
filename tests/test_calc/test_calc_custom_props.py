from __future__ import annotations
from typing import cast, TYPE_CHECKING
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from com.sun.star.beans import XPropertySet

from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.utils.info import Info

if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetDocument
    from com.sun.star.util import XCloseable
else:
    SpreadsheetDocument = object
    XCloseable = object


def test_get_props(copy_fix_calc, loader) -> None:
    doc_path = copy_fix_calc("custom_props.ods")
    doc = cast(SpreadsheetDocument, Calc.open_doc(fnm=str(doc_path), loader=loader))
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, idx=0)
        user_props = Info.get_user_defined_props(doc)
        assert user_props is not None
        ps = Lo.qi(XPropertySet, user_props, True)
        assert ps.getPropertyValue("PrintSheet") == 2.0

    finally:
        Lo.close(closeable=cast(XCloseable, doc), deliver_ownership=False)
