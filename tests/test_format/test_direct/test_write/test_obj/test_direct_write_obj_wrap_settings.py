from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.wrap import Settings, WrapTextMode
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def test_write(loader, formula_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)

        style = Settings(mode=WrapTextMode.THROUGH)

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(style,))

        f_style = Settings.from_obj(content)
        assert f_style.prop_mode == style.prop_mode

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
