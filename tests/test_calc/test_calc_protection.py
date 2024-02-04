from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.format.calc.direct.cell.cell_protection import CellProtection


def test_protection(loader) -> None:
    doc = Calc.create_doc()
    assert doc is not None, "Could not create new document"
    delay = 0
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)

        Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")

        cell = Calc.get_cell(sheet=sheet, cell_name="A1")

        assert Calc.is_cell_protected(cell=cell)

        cp = CellProtection(protected=False)
        cp.apply(cell)

        assert Calc.is_cell_protected(cell) == False

        cp = CellProtection(protected=True)
        cp.apply(cell)
        assert Calc.is_cell_protected(cell)

        assert Calc.is_sheet_protected(sheet=sheet) == False
        password = "1234"
        Calc.protect_sheet(sheet=sheet, password=password)
        assert Calc.is_sheet_protected(sheet=sheet)

        Calc.unprotect_sheet(sheet=sheet, password=password)
        assert Calc.is_sheet_protected(sheet=sheet) == False

        assert Calc.protect_sheet(sheet=sheet, password=password)
        assert Calc.is_sheet_protected(sheet=sheet)
        assert Calc.unprotect_sheet(sheet=sheet, password="abc") == False
        assert Calc.is_sheet_protected(sheet=sheet)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
