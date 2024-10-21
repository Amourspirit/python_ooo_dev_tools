from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.calc.modify.page.sheet import Printing, CalcStylePageKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


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

        style = Printing(header=False, grid=False, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Printing.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_header == style.prop_header
        assert f_style.prop_grid == style.prop_grid

        # ==========================================================
        style = Printing(comment=True, obj_img=False, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

        f_style = Printing.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_comment == style.prop_comment
        assert f_style.prop_obj_img == style.prop_obj_img

        # ==========================================================
        style = Printing(chart=False, drawing=False, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

        f_style = Printing.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_chart == style.prop_chart
        assert f_style.prop_style_name == style.prop_style_name

        # ==========================================================
        style = Printing(formula=True, zero_value=False, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

        f_style = Printing.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_formula == style.prop_formula
        assert f_style.prop_zero_value == style.prop_zero_value

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
