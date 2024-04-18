from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.style import Cell, StyleCellKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.color import CommonColor


def test_calc(loader) -> None:
    delay = 0

    doc = Calc.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        sheet = Calc.get_active_sheet()

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)

        style = Cell(name=StyleCellKind.ACCENT_1)
        xcell = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)
        style.apply(xcell)

        f_style = Cell.from_obj(xcell)
        assert f_style.prop_name == style.prop_name

        # ==============================================
        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)

        style = Cell().accent_2
        xcell = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)
        style.apply(xcell)

        f_style = Cell.from_obj(xcell)
        assert f_style.prop_name == style.prop_name

        # ==============================================
        data = TableHelper.make_2d_array(3, 3)
        rng_obj = RangeObj.from_range("A3:C5")
        Calc.set_array(values=data, sheet=sheet, range_obj=rng_obj)
        rng = Calc.get_cell_range(sheet=sheet, range_obj=rng_obj)

        style = Cell().good
        style.apply(rng)
        f_style = Cell.from_obj(rng)
        assert f_style.prop_name == style.prop_name

        style = Cell(name=StyleCellKind.DEFAULT)
        xprops = style.get_style_props()
        assert xprops is not None
        xprops.setPropertyValue("CellBackColor", CommonColor.CORAL)
        val = cast(int, xprops.getPropertyValue("CellBackColor"))
        assert val == CommonColor.CORAL

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
