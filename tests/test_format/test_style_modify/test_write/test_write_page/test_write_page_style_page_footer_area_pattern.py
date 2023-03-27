from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.footer import Footer, WriterStylePageKind
from ooodev.format.writer.modify.page.footer.area import Pattern, PresetPatternKind
from ooodev.format import Styler
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
        style = Pattern.from_preset(
            preset=PresetPatternKind.SHINGLE,
            style_name=footer_style.prop_style_name,
            style_family=footer_style.prop_style_family_name,
        )
        Styler.apply(doc, footer_style, style)
        # props = style.get_style_props(doc)

        f_style = Pattern.from_style(
            doc=doc, style_name=footer_style.prop_style_name, style_family=footer_style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_stretch == f_style.prop_inner.prop_stretch
        assert f_style.prop_inner.prop_tile == f_style.prop_inner.prop_tile

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
