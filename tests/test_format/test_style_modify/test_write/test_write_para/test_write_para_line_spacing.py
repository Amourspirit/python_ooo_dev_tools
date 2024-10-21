from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.para.indent_space import LineSpacing, ModeKind
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
        style = LineSpacing(mode=ModeKind.LINE_1_5, active_ln_spacing=True)
        style.apply(doc)
        props = style.get_style_props(doc)
        pp = cast("ParagraphProperties", props)
        assert style.prop_inner.prop_inner == pp.ParaLineSpacing

        f_style = LineSpacing.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_mode == ModeKind.LINE_1_5
        assert f_style.prop_inner.prop_active_ln_spacing == True

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
