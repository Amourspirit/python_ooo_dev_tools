from __future__ import annotations
from typing import cast, TYPE_CHECKING
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

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


def test_print_calc(copy_fix_calc, fix_printer_name, loader) -> None:
    if not fix_printer_name:
        pytest.skip("No printer configured")
    # because the test requires a printer, it is not run by default
    # test printing of a calc sheet direct to a printer
    # https://ask.libreoffice.org/t/direct-print-configured-cell-range/91933
    doc_path = copy_fix_calc("custom_props.ods")
    # custom_props.ods has a custom property "PrintSheet" set to 2
    doc = cast(SpreadsheetDocument, Calc.open_doc(fnm=str(doc_path), loader=loader))
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, idx=1)

        user_props = Info.get_user_defined_props(doc)
        assert user_props is not None
        ps = Lo.qi(XPropertySet, user_props, True)
        assert ps.getPropertyValue("PrintSheet") == 2.0
        sheet_idx = int(ps.getPropertyValue("PrintSheet")) - 1
        printer_name = fix_printer_name
        Calc.set_selected_addr(doc, sheet, "C6:G33")
        cell_rng = Calc.get_selected_addr(doc=doc)
        Calc.print_sheet(printer_name=printer_name, idx=sheet_idx, doc=doc, cr_addr=cell_rng)

    finally:
        Lo.close(closeable=cast(XCloseable, doc), deliver_ownership=False)
