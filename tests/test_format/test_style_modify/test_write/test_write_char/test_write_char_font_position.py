from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.char.font import FontPosition, Angle, FontScriptKind, CharSpacingKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
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

        style = FontPosition(script_kind=FontScriptKind.NORMAL, spacing=CharSpacingKind.LOOSE, rotation=Angle(90))
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("CharEscapementHeight") == 100
        assert props.getPropertyValue("CharRotation") == 900

        f_style = FontPosition.from_style(doc)
        assert f_style.prop_inner.prop_rotation == Angle(90)
        assert f_style.prop_inner.prop_script_kind == FontScriptKind.NORMAL
        assert f_style.prop_inner.prop_spacing == pytest.approx(CharSpacingKind.LOOSE.value, rel=1e2)
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
