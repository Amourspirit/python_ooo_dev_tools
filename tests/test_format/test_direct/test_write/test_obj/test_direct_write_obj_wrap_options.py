from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.wrap import (
    Options,
    Settings,
    WrapTextMode,
)
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def test_write(loader, formula_text) -> None:
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)

        style_settings = Settings(mode=WrapTextMode.THROUGH)

        # background can only be set to true when mode is Through
        style = Options(first=True, background=True)

        content = Write.add_formula(
            cursor=cursor,
            formula=formula_text,
            styles=(
                style_settings,
                style,
            ),
        )

        f_style = Options.from_obj(content)
        assert f_style.prop_first
        assert f_style.prop_first == style.prop_first
        assert f_style.prop_background
        assert f_style.prop_background == style.prop_background

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
