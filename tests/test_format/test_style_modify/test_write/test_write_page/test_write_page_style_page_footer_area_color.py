from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.footer import Footer, StylePageKind
from ooodev.format.writer.modify.page.footer.area import Color
from ooodev.format import Styler
from ooodev.utils.color import StandardColor
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

        footer_style = Footer(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
        )
        style = Color(
            color=StandardColor.GOLD_LIGHT2,
            style_name=footer_style.prop_style_name,
            style_family=footer_style.prop_style_family_name,
        )
        Styler.apply(doc, footer_style, style)
        # props = style.get_style_props(doc)

        footer_f_style = Footer.from_style(
            doc=doc, style_name=footer_style.prop_style_name, style_family=footer_style.prop_style_family_name
        )
        assert footer_f_style.prop_inner.prop_on == footer_style.prop_inner.prop_on
        assert footer_f_style.prop_inner.prop_shared_first == footer_style.prop_inner.prop_shared_first
        assert footer_f_style.prop_inner.prop_shared == footer_style.prop_inner.prop_shared
        assert footer_f_style.prop_inner.prop_height == pytest.approx(footer_style.prop_inner.prop_height, rel=1.0e2)
        assert footer_f_style.prop_inner.prop_margin_left == pytest.approx(
            footer_style.prop_inner.prop_margin_left, rel=1.0e2
        )
        assert footer_f_style.prop_inner.prop_margin_right == pytest.approx(
            footer_style.prop_inner.prop_margin_right, rel=1.0e2
        )

        f_style = Color.from_style(
            doc=doc, style_name=footer_style.prop_style_name, style_family=footer_style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_color == style.prop_inner.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
