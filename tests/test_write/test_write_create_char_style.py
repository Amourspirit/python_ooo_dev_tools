from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from typing import cast, TYPE_CHECKING
import uno

from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.info import Info
from ooodev.office.write import Write
from ooodev.units import UnitMM100
from ooodev.format.writer.direct.char.font import FontOnly as DirectFontOnly
from ooodev.format.writer.direct.char.highlight import Highlight as DirectHighlight
from ooodev.format.writer.modify.char.font import FontEffects as ModifyFontEffects, FontOnly as ModifyFontOnly
from ooodev.format.writer.modify.char.highlight import Highlight as ModifyHighlight
from ooodev.format.writer.style.char import Char as StyleChar
from ooodev.utils.color import StandardColor


def test_write(loader, para_text) -> None:
    # Test adding text frame with styles, as well as test getting the text frame from the document.

    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        style_name = "TestChar"

        # font 12 pt
        ft_size = UnitMM100.from_pt(14)
        font = DirectFontOnly(name=Info.get_font_general_name(), size=ft_size)

        hl = DirectHighlight(StandardColor.ORANGE_LIGHT3)

        _ = Write.create_style_char(text_doc=doc, style_name=style_name, styles=[font, hl])

        style_char = StyleChar(style_name)

        Write.style(3, 15, styles=(style_char,))

        f_font = ModifyFontOnly.from_style(doc=doc, style_name=style_name)

        f_highlight = ModifyHighlight.from_style(doc=doc, style_name=style_name)

        assert f_font.prop_inner.prop_name == Info.get_font_general_name()
        assert f_font.prop_inner.prop_size.get_value_mm100() in range(ft_size.value - 2, ft_size.value + 3)  # +- 2

        assert f_highlight.prop_inner.prop_color == StandardColor.ORANGE_LIGHT3

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
