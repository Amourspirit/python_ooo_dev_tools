from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.wrap import (
    Options,
    Settings,
    WrapTextMode,
)
from ooodev.format import Styler
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.units.unit_mm import UnitMM


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

        style_settings = Settings(mode=WrapTextMode.THROUGH)

        # background can only be set to true when mode is Through
        style = Options(first=True, background=True)

        frame = Write.add_text_frame(
            cursor=cursor,
            ypos=UnitMM(10.2),
            text=para_text,
            width=UnitMM(60),
            height=UnitMM(40),
            styles=(
                style_settings,
                style,
            ),
        )

        f_style = Options.from_obj(frame)
        assert f_style.prop_first
        assert f_style.prop_first == style.prop_first
        assert f_style.prop_background
        assert f_style.prop_background == style.prop_background

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
