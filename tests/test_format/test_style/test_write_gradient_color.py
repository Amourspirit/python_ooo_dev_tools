from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.style.writer.page.area.color import Color
from ooodev.format.style.writer.page.area.gradient import Gradient, StylePageKind, PresetKind
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

        style = Gradient.from_preset(PresetKind.SUNSHINE)
        style.apply(doc)

        obj = Gradient.from_obj(doc, style.prop_style_name)
        assert obj == style

        style = Gradient.from_preset(preset=PresetKind.MAHOGANY, style_name=StylePageKind.FIRST_PAGE)
        style.apply(doc)

        obj = Gradient.from_obj(doc, style.prop_style_name)
        assert obj == style

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
