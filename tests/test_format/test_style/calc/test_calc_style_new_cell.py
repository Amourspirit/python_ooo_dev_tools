from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.style import Cell, StyleCellKind

# from ooodev.format.calc.style import Page
from ooodev.format.calc.modify.cell.background import Color as BgColor
from ooodev.format.calc.modify.cell.font import FontEffects
from ooodev.office.calc import Calc
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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
        style_name = "Test Style"
        xstyle = Calc.create_cell_style(doc=doc, style_name=style_name)
        bg_color_style = BgColor(color=CommonColor.ROYAL_BLUE, style_name=style_name)
        fe_style = FontEffects(color=CommonColor.WHITE, style_name=style_name)
        Styler.apply(xstyle, bg_color_style, fe_style)

        style = Cell(name=style_name)
        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj, styles=[style])

        xcell = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)
        style.apply(xcell)

        # f_style = Cell.from_obj(xcell)
        f_style = Cell.from_obj(xstyle)
        assert f_style.prop_name == style.prop_name
        style_props = f_style.get_style_props()
        assert style_props is not None

        # pg = Page.from_obj(sheet)
        # pg_style_props = pg.get_style_props()
        # assert pg_style_props is not None

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
