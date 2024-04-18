from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area import Hatch, PresetHatchKind, WriterStylePageKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        presets = (
            PresetHatchKind.BLACK_0_DEGREES,
            PresetHatchKind.BLACK_180_DEGREES_CROSSED,
            PresetHatchKind.BLUE_45_DEGREES,
            PresetHatchKind.GREEN_90_DEGREES_TRIPLE,
            PresetHatchKind.RED_45_DEGREES_NEG_TRIPLE,
        )
        for preset in presets:

            style = Hatch.from_preset(preset=preset)
            style.apply(doc)
            # props = style.get_style_props(doc)
            # fp = cast("FillProperties", props)

            f_style = Hatch.from_style(doc, style.prop_style_name)
            assert f_style.prop_inner.prop_angle == style.prop_inner.prop_angle
            assert f_style.prop_inner.prop_bg_color == style.prop_inner.prop_bg_color
            assert f_style.prop_inner.prop_style == style.prop_inner.prop_style

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
