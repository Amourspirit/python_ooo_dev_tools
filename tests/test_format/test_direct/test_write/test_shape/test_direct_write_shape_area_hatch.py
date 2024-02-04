from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.shape.area import Hatch, HatchStyle, PresetHatchKind
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.office.draw import Draw
from ooodev.units.unit_mm import UnitMM


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        style = Hatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)

        page = Write.get_draw_page(doc)
        rs = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        style.apply(rs)
        page.add(rs)

        f_style = Hatch.from_obj(rs)
        assert f_style.prop_inner_color.prop_color == style.prop_inner_color.prop_color
        assert f_style.prop_inner_hatch == style.prop_inner_hatch

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
