from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.style import Frame, StyleFrameKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.units import UnitMM
from ooodev.utils.color import CommonColor


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        txt = "Hello"
        Write.append(cursor=cursor, text=txt)
        style = Frame(name=StyleFrameKind.FRAME)
        tf = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(20), text="World", width=UnitMM(40), height=UnitMM(40), styles=(style,)
        )

        f_style = Frame.from_obj(tf)
        assert f_style.prop_name == style.prop_name

        xprops = style.get_style_props()
        assert xprops is not None
        xprops.setPropertyValue("BackColor", CommonColor.CORAL)
        val = cast(int, xprops.getPropertyValue("BackColor"))
        assert val == CommonColor.CORAL

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
