from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.draw import Draw

# from ooodev.format.inner.direct.write.fill.area.pattern import Pattern, PresetPatternKind
from ooodev.format.draw.direct.area import Pattern, PresetPatternKind

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_draw_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        slide = Draw.get_slide(doc)
        rec = Draw.draw_rectangle(slide=slide, x=10, y=10, width=20, height=20)
        pattern = Pattern.from_preset(preset=PresetPatternKind.HORIZONTAL_BRICK)
        pattern.apply(rec)
        fp = cast("FillProperties", rec)
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == str(PresetPatternKind.HORIZONTAL_BRICK)
        assert fp.FillBitmap is not None

        cir = Draw.draw_circle(slide=slide, x=40, y=20, radius=10)
        pattern = Pattern.from_preset(preset=PresetPatternKind.DASHED_DOTTED_UPWARD_DIAGONAL)
        pattern.apply(cir)
        fp = cast("FillProperties", cir)
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == str(PresetPatternKind.DASHED_DOTTED_UPWARD_DIAGONAL)
        assert fp.FillBitmap is not None

        poly = Draw.draw_polygon(slide=slide, x=60, y=20, sides=5, radius=10)
        pattern = Pattern.from_preset(preset=PresetPatternKind.SHINGLE)
        pattern.apply(poly)
        fp = cast("FillProperties", poly)
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == str(PresetPatternKind.SHINGLE)
        assert fp.FillBitmap is not None

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
