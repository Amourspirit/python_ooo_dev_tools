from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area import Color, Gradient, StylePageKind, PresetGradientKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        style = Gradient.from_preset(PresetGradientKind.SUNSHINE)
        style.apply(doc)

        obj = Gradient.from_style(doc, style.prop_style_name)
        assert obj == style

        style = Gradient.from_preset(preset=PresetGradientKind.MAHOGANY, style_name=StylePageKind.FIRST_PAGE)
        style.apply(doc)

        obj = Gradient.from_style(doc, style.prop_style_name)
        assert obj == style

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
