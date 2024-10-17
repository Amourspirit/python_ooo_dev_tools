from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.para.drop_caps import DropCaps, StyleCharKind, StyleParaKind
from ooodev.format.inner.direct.structs.drop_cap_struct import DropCapStruct
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

        style = DropCaps(count=1, style=StyleCharKind.DROP_CAPS, style_name=StyleParaKind.CAPTION)
        style.apply(doc)
        props = style.get_style_props(doc)
        pp = cast("ParagraphProperties", props)
        assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, style.prop_inner._get_style_inst("drop_cap"))
        assert inner_dc == pp.DropCapFormat

        f_style = DropCaps.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_inner == style.prop_inner.prop_inner
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
