from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.modify.page.footer import Footer, CalcStylePageKind
from ooodev.format.calc.modify.page.footer.area import Img, PresetImageKind
from ooodev.format import Styler
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
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

        footer_style = Footer(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
            style_name=CalcStylePageKind.DEFAULT,
        )
        style = Img.from_preset(
            preset=PresetImageKind.COFFEE_BEANS,
            style_name=footer_style.prop_style_name,
            style_family=footer_style.prop_style_family_name,
        )
        Styler.apply(doc, footer_style, style)
        props = style.get_style_props(doc)

        graphic = props.getPropertyValue(style.prop_inner._props.back_graphic)
        assert graphic is not None

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
