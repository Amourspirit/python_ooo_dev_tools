from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.transparency import (
    Gradient,
    GradientStyle,
    Angle,
    Intensity,
    IntensityRange,
    Offset,
)
from ooodev.format.writer.direct.obj.area import Color
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def test_write(loader, formula_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)

        style = Gradient(style=GradientStyle.LINEAR, angle=Angle(45), grad_intensity=IntensityRange(0, 100))
        color_style = Color(color=StandardColor.GREEN_DARK3)

        content = Write.add_formula(
            cursor=cursor,
            formula=formula_text,
            styles=(
                color_style,
                style,
            ),
        )

        f_style = Gradient.from_obj(content)
        assert f_style.prop_inner == style.prop_inner

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
