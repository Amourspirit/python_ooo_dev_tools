from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.page.header import Header
from ooodev.format.writer.direct.page.header.area import Hatch, PresetHatchKind
from ooodev.format.writer.direct.char.font import Font
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info
from ooodev.units import UnitMM100


def test_write(loader, para_text) -> None:
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

        height = UnitMM100.from_mm(10.0)
        spacing = UnitMM100.from_mm(3.0)
        m_left = UnitMM100.from_mm(1.5)
        m_right = UnitMM100.from_mm(2.0)

        header = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=height,
            spacing=spacing,
            spacing_dyn=True,
            margin_left=m_left,
            margin_right=m_right,
        )
        style = Hatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)
        font = Font(b=True, color=StandardColor.RED_DARK3, size=16)
        Write.set_header(text_doc=doc, text="Header", styles=[header, font, style])

        props = Info.get_style_props(doc=doc, family_style_name="PageStyles", prop_set_nm="Standard")
        f_style = Hatch.from_obj(props)

        dist100 = style.prop_inner_hatch.prop_distance.get_value_mm100()
        assert f_style.prop_inner_hatch.prop_distance.get_value_mm100() in range(dist100 - 2, dist100 + 3)
        assert f_style.prop_inner_hatch.prop_angle == style.prop_inner_hatch.prop_angle
        assert f_style.prop_inner_hatch.prop_color == style.prop_inner_hatch.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
