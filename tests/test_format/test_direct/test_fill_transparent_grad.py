from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw
from ooodev.format.direct.fill.transparent.gradient import Gradient, GradientStyle, FillTransparentGrad


if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_draw_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = Draw.get_slide(doc)

        width = 36
        height = 36
        x = width / 2
        y = height / 2

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        ts = Gradient(angle=30, end_value=100)
        ts.apply(rec)
        fp = cast("FillProperties", rec)
        tp_grad = cast(FillTransparentGrad, ts._get_style_inst("fill_style"))
        assert tp_grad == fp.FillTransparenceGradient

        x += width
        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        ts = Gradient(style=GradientStyle.RECT, angle=25, end_value=100)
        ts.apply(rec)
        fp = cast("FillProperties", rec)
        tp_grad = cast(FillTransparentGrad, ts._get_style_inst("fill_style"))
        assert tp_grad == fp.FillTransparenceGradient

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
