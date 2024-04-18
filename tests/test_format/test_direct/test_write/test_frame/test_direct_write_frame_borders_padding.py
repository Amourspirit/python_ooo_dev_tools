from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.borders import Padding
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

        amt = 2.3
        style = Padding(all=amt)
        frame = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(10.2), text=para_text, width=UnitMM(60), height=UnitMM(40), styles=(style,)
        )

        f_style = Padding.from_obj(frame)
        assert f_style.prop_left.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_right.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_top.value == pytest.approx(amt, rel=1e-2)
        assert f_style.prop_bottom.value == pytest.approx(amt, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
