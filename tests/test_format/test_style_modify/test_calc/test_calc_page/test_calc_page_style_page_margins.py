from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.page.page import Margins, CalcStylePageKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc


def test_cacl(loader) -> None:
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

        style = Margins(left=10, right=10, top=18, bottom=18, style_name=CalcStylePageKind.DEFAULT)
        Styler.apply(doc, style)

        props = style.get_style_props(doc)
        for attrib in style.prop_inner._props:
            if attrib:
                val = cast(int, style.prop_inner._get(attrib))
                assert props.getPropertyValue(attrib) in range(val - 2, val + 3)

        f_style = Margins.from_style(doc, style.prop_style_name)

        for attrib in style.prop_inner._props:
            if attrib:
                val = cast(int, style.prop_inner._get(attrib))
                assert f_style.prop_inner._get(attrib) in range(val - 2, val + 3)
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
