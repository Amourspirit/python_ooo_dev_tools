from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast
import random

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.draw import Draw

# from ooodev.format.inner.direct.write.fill.area.hatch import Hatch, PresetHatchKind
from ooodev.format.inner.direct.structs.hatch_struct import HatchStruct
from ooodev.format.draw.direct.area import Hatch, PresetHatchKind
from ooodev.utils.color import StandardColor
from ooodev.utils.table_helper import TableHelper

from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.hatch_style import HatchStyle

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_presets_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_draw_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = Draw.get_slide(doc)
        presets = (
            PresetHatchKind.BLACK_0_DEGREES,
            PresetHatchKind.BLACK_90_DEGREES,
            PresetHatchKind.BLACK_180_DEGREES_CROSSED,
            PresetHatchKind.BLUE_45_DEGREES,
            PresetHatchKind.BLUE_45_DEGREES_NEG,
            PresetHatchKind.BLUE_45_DEGREES_CROSSED,
            PresetHatchKind.GREEN_30_DEGREES,
            PresetHatchKind.GREEN_60_DEGREES,
            PresetHatchKind.GREEN_90_DEGREES_TRIPLE,
            PresetHatchKind.RED_45_DEGREES,
            PresetHatchKind.RED_90_DEGREES_CROSSED,
            PresetHatchKind.RED_45_DEGREES_NEG_TRIPLE,
            PresetHatchKind.YELLOW_45_DEGREES,
            PresetHatchKind.YELLOW_45_DEGREES_CROSSED,
            PresetHatchKind.YELLOW_45_DEGREES_TRIPLE,
        )

        tbl = TableHelper.convert_1d_to_2d(presets, 5)
        width = 36
        height = 36
        y = 0
        for row, data_row in enumerate(tbl, start=1):
            x = width / 2
            if row == 1:
                y = height / 2
            else:
                y = (height / 2) + (height * (row - 1))

            for preset in data_row:
                rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
                hatch = Hatch.from_preset(preset)
                hatch.apply(rec)
                fp = cast("FillProperties", rec)
                assert fp.FillStyle == FillStyle.HATCH
                assert fp.FillBackground == False
                assert fp.FillColor == StandardColor.DEFAULT_BLUE
                assert fp.FillHatchName == str(preset)
                x += width

        offset = y + height
        for row, data_row in enumerate(tbl, start=1):
            x = width / 2
            if row == 1:
                y = (height / 2) + offset
            else:
                y = ((height / 2) + (height * (row - 1))) + offset

            for preset in data_row:
                color = StandardColor.get_random_color()
                bg_color = StandardColor.get_random_color()
                rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
                hatch = Hatch.from_preset(preset)
                hatch.prop_color = color
                hatch.prop_bg_color = bg_color
                hatch.apply(rec)
                hatch_struct = cast(HatchStruct, hatch._get_style("fill_hatch")[0])
                assert hatch_struct.prop_color == color
                fp = cast("FillProperties", rec)
                assert fp.FillStyle == FillStyle.HATCH
                assert fp.FillBackground == True
                assert fp.FillColor == bg_color
                assert hatch_struct == fp.FillHatch
                x += width

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_draw_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = Draw.get_slide(doc)

        width = 36
        height = 36
        x = width / 2
        y = height / 2

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        hatch = Hatch(style=HatchStyle.SINGLE, space=2.03, angle=45, color=StandardColor.BRICK_LIGHT2)
        hatch.apply(rec)
        fp = cast("FillProperties", rec)
        hatch_struct = cast(HatchStruct, hatch._get_style_inst("fill_hatch"))
        assert hatch_struct.prop_color == StandardColor.BRICK_LIGHT2
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == False
        assert fp.FillColor == StandardColor.DEFAULT_BLUE
        assert hatch_struct == fp.FillHatch

        x += width
        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        hatch = Hatch(
            style=HatchStyle.DOUBLE,
            space=3.2,
            angle=66,
            color=StandardColor.GOLD_LIGHT1,
            bg_color=StandardColor.BLUE_DARK3,
        )
        hatch.apply(rec)
        fp = cast("FillProperties", rec)
        hatch_struct = cast(HatchStruct, hatch._get_style_inst("fill_hatch"))
        assert hatch_struct.prop_color == StandardColor.GOLD_LIGHT1
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == True
        assert fp.FillColor == StandardColor.BLUE_DARK3
        assert hatch_struct == fp.FillHatch

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
