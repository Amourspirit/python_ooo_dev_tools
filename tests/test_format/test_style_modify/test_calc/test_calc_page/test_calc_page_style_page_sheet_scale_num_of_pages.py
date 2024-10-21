from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.calc.modify.page.sheet import ScaleNumOfPages, CalcStylePageKind
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

        style = ScaleNumOfPages(pages=3, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = ScaleNumOfPages.from_style(doc, style.prop_style_name)
        assert f_style.prop_pages == 3
        assert f_style.prop_pages == style.prop_pages

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
