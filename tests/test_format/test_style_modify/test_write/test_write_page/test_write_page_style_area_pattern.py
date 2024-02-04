from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area import Pattern, PresetPatternKind, WriterStylePageKind
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
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        presets = (
            PresetPatternKind.DASHED_DOTTED_UPWARD_DIAGONAL,
            PresetPatternKind.DIAGONAL_BRICK,
            PresetPatternKind.DIVOT,
            PresetPatternKind.SHINGLE,
            PresetPatternKind.ZIG_ZAG,
        )

        for preset in presets:

            style = Pattern.from_preset(preset=preset)
            style.apply(doc)
            # props = style.get_style_props(doc)
            # fp = cast("FillProperties", props)

            f_style = Pattern.from_style(doc, style.prop_style_name)
            assert f_style.prop_inner.prop_tile == style.prop_inner.prop_tile
            assert f_style.prop_inner.prop_stretch == style.prop_inner.prop_stretch

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
