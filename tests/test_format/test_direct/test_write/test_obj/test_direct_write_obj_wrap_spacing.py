from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.wrap import Spacing
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.units.unit_mm100 import UnitMM100


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

        style = Spacing(all=2.8)

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(style,))

        f_style = Spacing.from_obj(content)
        assert f_style.prop_left.value == pytest.approx(style.prop_left.value, rel=1e-2)
        assert f_style.prop_right.value == pytest.approx(style.prop_right.value, rel=1e-2)
        assert f_style.prop_top.value == pytest.approx(style.prop_top.value, rel=1e-2)
        assert f_style.prop_bottom.value == pytest.approx(style.prop_bottom.value, rel=1e-2)

        style = Spacing(all=UnitMM100(280))
        style.apply(content)
        f_style = Spacing.from_obj(content)
        assert f_style.prop_left.value == pytest.approx(style.prop_left.value, rel=1e-2)
        assert f_style.prop_right.value == pytest.approx(style.prop_right.value, rel=1e-2)
        assert f_style.prop_top.value == pytest.approx(style.prop_top.value, rel=1e-2)
        assert f_style.prop_bottom.value == pytest.approx(style.prop_bottom.value, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
