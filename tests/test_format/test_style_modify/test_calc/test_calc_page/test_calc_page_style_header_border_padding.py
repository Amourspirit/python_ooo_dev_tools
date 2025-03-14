from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.calc.modify.page.header import Header, CalcStylePageKind
from ooodev.format.calc.modify.page.header.borders import Padding
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.format import Styler
from ooodev.office.calc import Calc
from ooodev.units.unit_mm100 import UnitMM100


def test_calc(loader) -> None:
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

        header_style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            margin_left=1.5,
            margin_right=2.0,
        )
        amt = UnitMM100.from_mm(7.3)
        style = Padding(padding_all=amt, style_name=CalcStylePageKind.DEFAULT)
        Styler.apply(doc, header_style, style)
        props = style.get_style_props(doc)

        rng = range(amt.value - 2, amt.value + 3)  # +- 2
        assert props.getPropertyValue("HeaderLeftBorderDistance") in rng

        f_style = Padding.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_left.get_value_mm100() in rng
        assert f_style.prop_inner.prop_right.get_value_mm100() in rng
        assert f_style.prop_inner.prop_top.get_value_mm100() in rng
        assert f_style.prop_inner.prop_bottom.get_value_mm100() in rng

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
