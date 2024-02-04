from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.image.options import Names
from ooodev.format.writer.direct.image.options import Protect
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

        style_name = Names(name="formula", desc="Just a test Formula", alt="A real Formula")
        style = Protect(size=True, position=True, content=True)

        content = Write.add_formula(
            cursor=cursor,
            formula=formula_text,
            styles=(
                style_name,
                style,
            ),
        )

        f_style = Protect.from_obj(content)
        assert f_style.prop_size == style.prop_size
        assert f_style.prop_position == style.prop_position
        assert f_style.prop_content == style.prop_content

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
