from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.frame.area import StyleFrameKind, Hatch, InnerHatch, HatchStyle, PresetHatchKind
from ooodev.utils.color import StandardColor
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

        style = Hatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES, style_name=StyleFrameKind.FRAME)

        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Hatch.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_inner_color.prop_color == style.prop_inner.prop_inner_color.prop_color
        assert f_style.prop_inner.prop_inner_hatch == style.prop_inner.prop_inner_hatch

        inner_hatch = InnerHatch.from_preset(preset=PresetHatchKind.RED_90_DEGREES_CROSSED)
        style.prop_inner = inner_hatch
        style.apply(doc)

        f_style = Hatch.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_inner_color.prop_color == style.prop_inner.prop_inner_color.prop_color
        assert f_style.prop_inner.prop_inner_hatch == style.prop_inner.prop_inner_hatch

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
