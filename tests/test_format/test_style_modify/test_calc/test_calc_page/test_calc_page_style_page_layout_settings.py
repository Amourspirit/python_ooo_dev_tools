from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format import Styler
from ooodev.format.calc.modify.page.page import LayoutSettings, PageStyleLayout, NumberingTypeEnum
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

        style = LayoutSettings(
            layout=PageStyleLayout.MIRRORED,
            numbers=NumberingTypeEnum.CHARS_UPPER_LETTER,
            align_hori=True,
            align_vert=True,
        )
        Styler.apply(doc, style)
        # props = style.get_style_props(doc)

        f_style = LayoutSettings.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_layout == style.prop_inner.prop_layout
        assert f_style.prop_inner.prop_numbers == style.prop_inner.prop_numbers
        assert f_style.prop_inner.prop_align_hori == style.prop_inner.prop_align_hori
        assert f_style.prop_inner.prop_align_vert == style.prop_inner.prop_align_vert

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
