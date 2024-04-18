"""Test for ooodev.format.writer.direct.shape.area.Color"""

# pylint: disable=no-member
# pylint: disable=unused-import
# pylint: disable=unused-argument
# pylint: disable=wrong-import-order
# pylint: disable=wrong-import-position
# pylint: disable=missing-function-docstring
# pylint: disable=missing-class-docstring
# pylint: disable=invalid-name
from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.shape.area import Color
from ooodev.utils.color import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.office.draw import Draw


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        style = Color(color=StandardColor.GREEN_LIGHT2)

        page = Write.get_draw_page(doc)
        rs = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        style.apply(rs)
        page.add(rs)

        f_style = Color.from_obj(rs)
        assert f_style.prop_color == style.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
