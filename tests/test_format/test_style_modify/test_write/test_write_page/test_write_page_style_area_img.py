from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area import Img, PresetImageKind, WriterStylePageKind
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
            PresetImageKind.BATHROOM_TILES,
            PresetImageKind.BRICK_WALL,
            PresetImageKind.COFFEE_BEANS,
            PresetImageKind.INVOICE_PAPER,
        )

        for preset in presets:
            style = Img.from_preset(preset=preset)
            style.apply(doc)
            # props = style.get_style_props(doc)

            f_style = Img.from_style(doc, style.prop_style_name)
            assert f_style.prop_inner.prop_size == style.prop_inner.prop_size
            assert f_style.prop_inner.prop_mode == style.prop_inner.prop_mode
            assert f_style.prop_inner.prop_pos_offset == style.prop_inner.prop_pos_offset
            assert f_style.prop_inner.prop_position == style.prop_inner.prop_position

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
