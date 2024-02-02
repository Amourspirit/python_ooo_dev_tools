from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.modify.area import Img, PresetImageKind, SizeMM

from ooo.dyn.drawing.fill_style import FillStyle


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = DrawDoc(Draw.create_draw_doc())
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = doc.get_slide(idx=0)

        width = 100
        height = 100
        x = int(width / 2)
        y = int(height / 2)

        style = Img.from_preset(preset=PresetImageKind.FLORAL)
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        props = style.get_style_props(doc.component)
        assert props.getPropertyValue("FillStyle") == FillStyle.BITMAP
        assert props.getPropertyValue("FillBitmapName") == str(PresetImageKind.FLORAL)

        f_style = Img.from_style(doc.component)
        point = PresetImageKind.FLORAL._get_point()
        x_lst = [(point.x - 2) + i for i in range(5)]  # plus or minus 2
        y_lst = [(point.y - 2) + i for i in range(5)]  # plus or minus 2
        assert f_style.prop_inner.prop_is_size_mm
        size = f_style.prop_inner.prop_size
        assert isinstance(size, SizeMM)
        assert round(size.width * 100) in x_lst
        assert round(size.height * 100) in y_lst

        Lo.delay(delay)
    finally:
        doc.close_doc()
