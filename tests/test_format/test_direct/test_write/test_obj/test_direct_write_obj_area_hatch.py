from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.area import Hatch, HatchStyle, PresetHatchKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
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

        style = Hatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(style,))

        f_style = Hatch.from_obj(content)
        assert f_style.prop_inner_color.prop_color == style.prop_inner_color.prop_color
        assert f_style.prop_inner_hatch == style.prop_inner_hatch

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
