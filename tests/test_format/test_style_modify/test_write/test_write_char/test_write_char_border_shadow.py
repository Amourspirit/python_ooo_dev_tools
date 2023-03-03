from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.char.borders import Shadow, ShadowFormat, ShadowLocation
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

        style = Shadow(location=ShadowLocation.BOTTOM_RIGHT, width=2.0)
        style.apply(doc)
        props = style.get_style_props(doc)
        struct = style.prop_inner.get_uno_struct()
        p_struct = cast(ShadowFormat, props.getPropertyValue("CharShadowFormat"))
        assert struct.Color == p_struct.Color
        assert struct.Location == p_struct.Location

        f_style = Shadow.from_style(doc)
        assert f_style.prop_inner.prop_location == ShadowLocation.BOTTOM_RIGHT
        assert f_style.prop_inner.prop_width.value == pytest.approx(2.0, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
