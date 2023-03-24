from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.modify.page.header import Header, CalcStylePageKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc
from ooodev.utils.data_type.unit_mm100 import UnitMM100


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
        height100 = UnitMM100.from_mm(7)
        spacing100 = UnitMM100.from_mm(3)
        left100 = UnitMM100.from_mm(1.5)
        right100 = UnitMM100.from_mm(2)

        style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=height100,
            spacing=spacing100,
            spacing_dyn=True,
            margin_left=left100,
            margin_right=right100,
            height_auto=False,
        )
        style.apply(doc)
        props = style.get_style_props(doc)

        f_style = Header.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_on == style.prop_inner.prop_on
        assert f_style.prop_inner.prop_shared_first == style.prop_inner.prop_shared_first
        assert f_style.prop_inner.prop_shared == style.prop_inner.prop_shared
        assert f_style.prop_inner.prop_height.get_value_mm100() in range(
            height100.value - 2, height100.value + 3
        )  # +- 2
        assert f_style.prop_inner.prop_margin_left.get_value_mm100() in range(
            left100.value - 2, left100.value + 3
        )  # +- 2
        assert f_style.prop_inner.prop_margin_right.get_value_mm100() in range(
            right100.value - 2, right100.value + 3
        )  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
