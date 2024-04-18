from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.borders import Side, Sides, BorderLineKind, LineSize
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import StandardColor


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

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED_DARK3, width=LineSize.MEDIUM)
        style = Sides(all=side)

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(style,))

        f_style = Sides.from_obj(content)
        f_side = f_style.prop_left
        assert f_side.prop_color == side.prop_color
        assert f_side.prop_width.value == pytest.approx(side.prop_width.value, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
