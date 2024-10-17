from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format import Styler
from ooodev.format.calc.modify.cell.cell_protection import CellProtection, StyleCellKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


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

        style = CellProtection(
            hide_all=False, hide_formula=False, protected=True, hide_print=True, style_name=StyleCellKind.DEFAULT
        )
        Styler.apply(doc, style)

        f_style = CellProtection.from_style(doc)
        assert f_style.prop_inner.prop_hide_all == style.prop_inner.prop_hide_all
        assert f_style.prop_inner.prop_hide_formula == style.prop_inner.prop_hide_formula
        assert f_style.prop_inner.prop_protected == style.prop_inner.prop_protected
        assert f_style.prop_inner.prop_hide_print == style.prop_inner.prop_hide_print

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
