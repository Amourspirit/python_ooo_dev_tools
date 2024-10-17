from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format import Styler
from ooodev.format.calc.modify.cell.font import FontOnly, StyleCellKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.units.unit_mm100 import UnitMM100


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
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

        font_size = UnitMM100.from_pt(12.0)
        style = FontOnly(name="Liberation Sans", size=font_size, font_style="Bold", style_name=StyleCellKind.DEFAULT)
        Styler.apply(doc, style)

        f_style = FontOnly.from_style(doc)
        assert f_style.prop_inner.prop_name == style.prop_inner.prop_name
        assert f_style.prop_inner.prop_size.get_value_mm100() in range(
            font_size.value - 2, font_size.value + 3
        )  # +- 2
        assert f_style.prop_inner.prop_style_name == style.prop_inner.prop_style_name

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
