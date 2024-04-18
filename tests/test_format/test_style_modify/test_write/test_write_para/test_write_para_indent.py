from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.indent_space import Indent, StyleParaKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        amt = 6.0
        style = Indent(first=amt)
        style.apply(doc)
        props = style.get_style_props(doc)
        pp = cast("ParagraphProperties", props)
        assert pp.ParaFirstLineIndent in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2

        f_style = Indent.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_first.value == pytest.approx(amt)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
