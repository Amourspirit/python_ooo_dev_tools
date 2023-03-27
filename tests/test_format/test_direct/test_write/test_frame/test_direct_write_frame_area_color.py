from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.area import Color
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
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

        style = Color(color=StandardColor.GREEN_LIGHT2)

        frame = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(10.2), text=para_text, width=UnitMM(60), height=UnitMM(40), styles=(style,)
        )

        f_style = Color.from_obj(frame)
        assert f_style.prop_color == style.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
