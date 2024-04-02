from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.borders import Padding
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
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        style = Padding(all=3.0)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("LeftBorderDistance") in (298, 299, 300, 301, 302)

        f_style = Padding.from_style(doc)
        assert f_style.prop_inner.prop_left.value == pytest.approx(3.0, rel=1e-2)
        assert f_style.prop_inner.prop_top.value == pytest.approx(3.0, rel=1e-2)
        assert f_style.prop_inner.prop_right.value == pytest.approx(3.0, rel=1e-2)
        assert f_style.prop_inner.prop_bottom.value == pytest.approx(3.0, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
