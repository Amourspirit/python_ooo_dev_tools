from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.area import Img, PresetImageKind, SizeMM
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

        style = Img.from_preset(preset=PresetImageKind.COLORFUL_PEBBLES)

        frame = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(10.2), text=para_text, width=UnitMM(60), height=UnitMM(40), styles=(style,)
        )

        f_style = Img.from_obj(frame)
        point = PresetImageKind.COLORFUL_PEBBLES._get_point()
        xlst = [(point.x - 2) + i for i in range(5)]  # plus or minus 2
        ylst = [(point.y - 2) + i for i in range(5)]  # plus or minus 2
        assert f_style.prop_is_size_mm
        size = f_style.prop_size
        assert isinstance(size, SizeMM)
        assert round(size.width * 100) in xlst
        assert round(size.height * 100) in ylst

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
