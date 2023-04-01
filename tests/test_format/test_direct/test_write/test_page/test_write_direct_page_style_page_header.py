from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.page.header import Header
from ooodev.format.writer.direct.char.font import Font
from ooodev.format.writer.direct.para.alignment import Alignment, ParagraphAdjust
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

        style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=height,
            spacing=spacing,
            spacing_dyn=True,
            margin_left=m_left,
            margin_right=m_right,
        )
        align = Alignment(align=ParagraphAdjust.CENTER)
        font = Font(b=True, color=StandardColor.RED_DARK3, size=16)
        Write.set_header(text_doc=doc, text="Header", styles=[style, font, align])
        # props = style.get_style_props(doc)

        props = Info.get_style_props(doc=doc, family_style_name="PageStyles", prop_set_nm="Standard")
        val = props.getPropertyValue(style._props.shared_first)
        assert val
        val = props.getPropertyValue(style._props.shared)
        assert val
        # height is somewhat dynamic, so we can't test it.
        # val = cast(int, props.getPropertyValue(style._props.height))
        # assert val in range(height.value - 2, height.value + 3)
        val = cast(int, props.getPropertyValue(style._props.spacing))
        assert val in range(spacing.value - 2, spacing.value + 3)
        val = cast(int, props.getPropertyValue(style._props.margin_left))
        assert val in range(m_left.value - 2, m_left.value + 3)
        val = cast(int, props.getPropertyValue(style._props.margin_right))
        assert val in range(m_right.value - 2, m_right.value + 3)
        val = props.getPropertyValue(style._props.spacing_dyn)
        assert val

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
