from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.frame.borders import StyleFrameKind, Padding, InnerPadding
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
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        amt = 15.3
        style = Padding(all=amt, style_name=StyleFrameKind.FRAME)
        style.apply(doc)
        props = style.get_style_props(doc)

        border = round(amt * 100)
        assert props.getPropertyValue("LeftBorderDistance") in range(border - 2, border + 3)

        f_style = Padding.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_left == pytest.approx(amt, rel=1e2)
        assert f_style.prop_inner.prop_right == pytest.approx(amt, rel=1e2)
        assert f_style.prop_inner.prop_top == pytest.approx(amt, rel=1e2)
        assert f_style.prop_inner.prop_bottom == pytest.approx(amt, rel=1e2)

        amt = 12.4
        inner = InnerPadding(all=amt)
        style.prop_inner = inner

        style.apply(doc)
        props = style.get_style_props(doc)

        border = round(amt * 100)
        assert props.getPropertyValue("LeftBorderDistance") in range(border - 2, border + 3)

        f_style = Padding.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_left == pytest.approx(amt, rel=1e2)
        assert f_style.prop_inner.prop_right == pytest.approx(amt, rel=1e2)
        assert f_style.prop_inner.prop_top == pytest.approx(amt, rel=1e2)
        assert f_style.prop_inner.prop_bottom == pytest.approx(amt, rel=1e2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
