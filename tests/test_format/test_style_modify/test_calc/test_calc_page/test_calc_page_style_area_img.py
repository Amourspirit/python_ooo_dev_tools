from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.calc.modify.page.area import Img, PresetImageKind, CalcStylePageKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_calc(loader) -> None:
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

        presets = (
            PresetImageKind.CARDBOARD,
            PresetImageKind.CONCRETE,
            PresetImageKind.ICE_LIGHT,
        )

        for preset in presets:
            style = Img.from_preset(preset=preset, style_name=CalcStylePageKind.DEFAULT)
            style.apply(doc)
            props = style.get_style_props(doc)
            transparent = props.getPropertyValue(style.prop_inner._props.transparent)
            assert transparent
            graphic = props.getPropertyValue(style.prop_inner._props.back_graphic)
            assert graphic is not None

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
