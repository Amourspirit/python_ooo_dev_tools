from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.modify.page.sheet import Order, CalcStylePageKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc


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

        style = Order(top_btm=False, first_pg=0, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Order.from_style(doc, style.prop_style_name)
        assert f_style.prop_top_btm == style.prop_top_btm
        assert f_style.prop_first_pg == style.prop_first_pg

        # ============================================
        style = Order(top_btm=True, first_pg=2, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Order.from_style(doc, style.prop_style_name)
        assert f_style.prop_top_btm == style.prop_top_btm
        assert f_style.prop_first_pg == style.prop_first_pg

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
