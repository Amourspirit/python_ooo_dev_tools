from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.shape.shadow import Shadow, ShadowLocationKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.office.draw import Draw
from ooodev.utils.color import StandardColor


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        style = Shadow(
            use_shadow=True,
            location=ShadowLocationKind.BOTTOM_RIGHT,
            color=StandardColor.RED_LIGHT2,
            distance=5.0,
            blur=2,
            transparency=20,
        )

        page = Write.get_draw_page(doc)
        rs = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        style.apply(rs)
        page.add(rs)

        f_style = Shadow.from_obj(rs)
        assert f_style.prop_use_shadow == style.prop_use_shadow
        assert round(f_style.prop_blur.value) == round(style.prop_blur.value)
        assert f_style.prop_color == style.prop_color
        assert f_style.prop_distance.value == pytest.approx(style.prop_distance.value, rel=1e-2)
        assert f_style.prop_transparency == style.prop_transparency

        kinds = (
            ShadowLocationKind.TOP_LEFT,
            ShadowLocationKind.TOP,
            ShadowLocationKind.TOP_RIGHT,
            ShadowLocationKind.RIGHT,
            ShadowLocationKind.BOTTOM_RIGHT,
            ShadowLocationKind.BOTTOM,
            ShadowLocationKind.BOTTOM_LEFT,
            ShadowLocationKind.LEFT,
        )
        for kind in kinds:
            kind_style = style.fmt_location(kind).fmt_color(StandardColor.get_random_color())
            kind_style.apply(rs)

            f_style = Shadow.from_obj(rs)
            assert f_style.prop_use_shadow == kind_style.prop_use_shadow
            assert round(f_style.prop_blur.value) == round(kind_style.prop_blur.value)
            assert f_style.prop_color == kind_style.prop_color
            assert f_style.prop_distance.value == pytest.approx(kind_style.prop_distance.value, rel=1e-2)
            assert f_style.prop_transparency == kind_style.prop_transparency

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
