from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.page.header import Header
from ooodev.format.writer.modify.page.header.borders import Padding
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.format import Styler
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

        header_style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
        )
        amt = 7.3
        style = Padding(padding_all=amt)
        Styler.apply(doc, header_style, style)
        props = style.get_style_props(doc)

        border = round(amt * 100)
        assert props.getPropertyValue("HeaderLeftBorderDistance") in range(border - 2, border + 3)

        f_style = Padding.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_left.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_inner.prop_right.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_inner.prop_top.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_inner.prop_bottom.value == pytest.approx(amt, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
