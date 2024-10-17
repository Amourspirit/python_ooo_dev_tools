from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.image.options import Names
from ooodev.format.writer.direct.image.options import Properties
from ooodev.gui.gui import GUI
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

        style = Properties(printable=False)

        content = Write.add_formula(
            cursor=cursor,
            formula=formula_text,
            styles=(
                style_name,
                style,
            ),
        )

        f_style = Properties.from_obj(content)
        assert f_style.prop_printable == style.prop_printable

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
