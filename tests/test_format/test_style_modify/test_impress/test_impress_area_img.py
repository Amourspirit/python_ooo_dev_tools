from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw
from ooodev.format.impress.modify.area import (
    Img,
    PresetImageKind,
    ImgStyleKind,
    SizeMM,
    SizePercent,
    Offset,
    OffsetColumn,
    OffsetRow,
    RectanglePoint,
)

from ooo.dyn.drawing.fill_style import FillStyle


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_impress_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = Draw.get_slide(doc)

        width = 100
        height = 100
        x = width / 2
        y = height / 2

        style = Img.from_preset(preset=PresetImageKind.FLORAL)
        style.apply(doc)

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("FillStyle") == FillStyle.BITMAP
        assert props.getPropertyValue("FillBitmapName") == str(PresetImageKind.FLORAL)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
