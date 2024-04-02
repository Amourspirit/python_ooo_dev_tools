from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.numbers import Numbers, NumberFormatEnum, NumberFormatIndexEnum, StyleCellKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


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
        style = Numbers(num_format_index=NumberFormatIndexEnum.CURRENCY_1000INT_RED, style_name=StyleCellKind.DEFAULT)
        Calc.set_val(value=-123.45, sheet=sheet, cell_obj=cell_obj)

        Styler.apply(doc, style)

        f_style = Numbers.from_style(doc)
        assert f_style.prop_inner.prop_format_key == style.prop_inner.prop_format_key

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
