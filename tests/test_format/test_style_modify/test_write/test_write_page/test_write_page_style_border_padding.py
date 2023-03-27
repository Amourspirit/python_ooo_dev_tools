from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.borders import Padding, WriterStylePageKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import StandardColor
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

        amt = 15.3
        style = Padding(all=amt)
        style.apply(doc)
        props = style.get_style_props(doc)

        border = round(amt * 100)
        assert props.getPropertyValue("LeftBorderDistance") in range(border - 2, border + 3)

        f_style = Padding.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_left.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_inner.prop_right.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_inner.prop_top.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_inner.prop_bottom.value == pytest.approx(amt, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
