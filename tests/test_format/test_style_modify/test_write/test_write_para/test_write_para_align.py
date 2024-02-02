from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.alignment import Alignment, LastLineKind, ParagraphAdjust
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


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

        style = Alignment(align=ParagraphAdjust.BLOCK, align_last=LastLineKind.CENTER)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("ParaLastLineAdjust") == LastLineKind.CENTER.value
        assert props.getPropertyValue("ParaAdjust") == 2

        f_style = Alignment.from_style(doc)
        assert f_style.prop_inner.prop_align == ParagraphAdjust.BLOCK
        assert f_style.prop_inner.prop_align_last == LastLineKind.CENTER
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
