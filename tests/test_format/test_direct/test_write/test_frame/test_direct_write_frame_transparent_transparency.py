from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.frame.transparency import Transparency, Intensity
from ooodev.format.writer.direct.frame.area import Color
from ooodev.utils.color import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.units.unit_mm import UnitMM


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

        style = style = Transparency(value=Intensity(52))
        color_style = Color(color=StandardColor.GOLD_DARK4)

        frame = Write.add_text_frame(
            cursor=cursor,
            ypos=UnitMM(10.2),
            text=para_text,
            width=UnitMM(60),
            height=UnitMM(40),
            styles=(
                color_style,
                style,
            ),
        )

        f_style = Transparency.from_obj(frame)
        assert f_style.prop_value == style.prop_value

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
