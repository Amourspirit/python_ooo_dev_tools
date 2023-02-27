from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.options import Properties, TextDirectionKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.data_type.unit_mm import UnitMM


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

        style = Properties(editable=True, printable=True, txt_direction=TextDirectionKind.LR_TB)

        frame = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(10.2), text=para_text, width=UnitMM(60), height=UnitMM(40), styles=(style,)
        )

        f_style = Properties.from_obj(frame)
        assert f_style.prop_editable == style.prop_editable
        assert f_style.prop_printable == style.prop_printable
        assert f_style.prop_txt_direction == style.prop_txt_direction

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
