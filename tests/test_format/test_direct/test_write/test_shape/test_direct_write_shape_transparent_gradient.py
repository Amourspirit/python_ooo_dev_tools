from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.shape.transparency import (
    Gradient,
    GradientStyle,
    Angle,
    Intensity,
    IntensityRange,
    Offset,
)
from ooodev.format.writer.direct.frame.area import Img, PresetImageKind
from ooodev.format import Styler
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.draw import Draw
from ooodev.office.write import Write


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        style = Gradient(style=GradientStyle.LINEAR, angle=Angle(45), grad_intensity=IntensityRange(0, 100))
        img_style = Img.from_preset(preset=PresetImageKind.COLOR_STRIPES)

        page = Write.get_draw_page(doc)
        rs = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        page.add(rs)
        Styler.apply(rs, img_style, style)

        f_style = Gradient.from_obj(rs)
        assert f_style.prop_inner == style.prop_inner

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
