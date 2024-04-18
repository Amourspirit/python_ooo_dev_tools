from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from typing import cast, TYPE_CHECKING
import uno

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.info import Info
from ooodev.office.write import Write
from ooodev.units import UnitMM100
from ooodev.format.writer.direct.char.font import FontOnly as DirectFontOnly
from ooodev.format.writer.direct.para.indent_space import (
    Spacing as DirectSpacing,
    LineSpacing as DirectLineSpacing,
    ModeKind,
)
from ooodev.format.writer.modify.para.font import FontOnly as ModifyFontOnly
from ooodev.format.writer.modify.para.indent_space import Spacing as ModifySpacing, LineSpacing as ModifyLineSpacing
from ooodev.format.writer.style.para import Para as StylePara


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

        style_name = "TestABC"

        # font 12 pt
        ft_size = UnitMM100.from_pt(12)
        font = DirectFontOnly(name=Info.get_font_general_name(), size=ft_size)

        # spacing below paragraphs
        spc_size = UnitMM100.from_mm(4)
        spc = DirectSpacing(below=spc_size)

        # paragraph line spacing
        ln_spc_size = UnitMM100.from_mm(6)
        ln_spc = DirectLineSpacing(mode=ModeKind.FIXED, value=ln_spc_size)

        _ = Write.create_style_para(text_doc=doc, style_name=style_name, styles=[font, spc, ln_spc])

        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text, styles=(StylePara(style_name),))

        f_font = ModifyFontOnly.from_style(doc=doc, style_name=style_name)
        f_spacing = ModifySpacing.from_style(doc=doc, style_name=style_name)
        f_ln_spacing = ModifyLineSpacing.from_style(doc=doc, style_name=style_name)

        assert f_font.prop_inner.prop_name == Info.get_font_general_name()
        assert f_font.prop_inner.prop_size.get_value_mm100() in range(ft_size.value - 2, ft_size.value + 3)  # +- 2

        assert f_spacing.prop_inner.prop_below.get_value_mm100() in range(spc_size.value - 2, spc_size.value + 3)

        assert f_ln_spacing.prop_inner.prop_mode == ModeKind.FIXED
        assert f_ln_spacing.prop_inner.prop_value in range(ln_spc_size.value - 2, ln_spc_size.value + 3)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
