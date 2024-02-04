from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.font import (
    FontEffects,
    FontUnderlineEnum,
    FontStrikeoutEnum,
    FontLine,
    StyleCellKind,
)
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.utils.color import StandardColor


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

        style = FontEffects(
            color=StandardColor.BLUE,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.RED),
            strike=FontStrikeoutEnum.SINGLE,
            style_name=StyleCellKind.DEFAULT,
        )
        Styler.apply(doc, style)

        f_style = FontEffects.from_style(doc)
        assert f_style.prop_inner.prop_color == StandardColor.BLUE
        assert f_style.prop_inner.prop_underline.color == StandardColor.RED
        assert f_style.prop_inner.prop_underline.line == FontUnderlineEnum.SINGLE
        assert f_style.prop_inner.prop_strike == FontStrikeoutEnum.SINGLE

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
