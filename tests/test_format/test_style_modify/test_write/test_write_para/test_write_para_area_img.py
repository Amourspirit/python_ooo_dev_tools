from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.area import StyleParaKind, Img, PresetImageKind, SizeMM
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write

from ooo.dyn.drawing.fill_style import FillStyle


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

        style = Img.from_preset(PresetImageKind.COLORFUL_PEBBLES, StyleParaKind.STANDARD)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("FillStyle") == FillStyle.BITMAP
        assert props.getPropertyValue("FillBitmapName") == str(PresetImageKind.COLORFUL_PEBBLES)
        # assert props.getPropertyValue("FillGradientName") == str(PresetGradientKind.MAHOGANY)

        f_style = Img.from_style(doc)
        point = PresetImageKind.COLORFUL_PEBBLES._get_point()
        xlst = [(point.x - 2) + i for i in range(5)]  # plus or minus 2
        ylst = [(point.y - 2) + i for i in range(5)]  # plus or minus 2
        assert f_style.prop_inner.prop_is_size_mm
        size = f_style.prop_inner.prop_size
        assert isinstance(size, SizeMM)
        assert round(size.width * 100) in xlst
        assert round(size.height * 100) in ylst

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
