from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.modify.page.borders import (
    Sides,
    Side,
    LineSize,
    CalcStylePageKind,
    BorderLineKind,
)
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import StandardColor
from ooodev.office.calc import Calc
from ooodev.utils.data_type.unit_mm100 import UnitMM100


def test_write(loader) -> None:
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

        amt = UnitMM100.from_pt(float(LineSize.MEDIUM))
        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED_DARK3, width=amt)

        style = Sides(all=side, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Sides.from_style(doc, style.prop_style_name)
        f_side = f_style.prop_inner.prop_left
        assert f_side.prop_color == side.prop_color
        assert f_side.prop_width.get_value_mm100() in range(amt.value - 2, amt.value + 3)  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
