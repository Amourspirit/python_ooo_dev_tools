from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.obj.borders import ShadowLocation, Shadow
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import StandardColor


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

        style = Shadow(location=ShadowLocation.BOTTOM_RIGHT, color=StandardColor.GRAY_LIGHT1, width=2.4)

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(style,))

        f_style = Shadow.from_obj(content)
        assert f_style.prop_width.value == pytest.approx(style.prop_width.value, rel=1e-2)
        assert f_style.prop_color == StandardColor.GRAY_LIGHT1
        assert f_style.prop_location == ShadowLocation.BOTTOM_RIGHT

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
