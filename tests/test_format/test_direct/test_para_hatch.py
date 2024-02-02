from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write

# from ooodev.format.inner.direct.write.para.area.hatch import Hatch, PresetHatchKind, HatchStyle, Angle
from ooodev.format.writer.direct.para.area import Hatch, PresetHatchKind, HatchStyle, Angle
from ooodev.utils.color import StandardColor

from ooo.dyn.drawing.fill_style import FillStyle

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service
    from com.sun.star.style import ParagraphProperties  # service


def test_write_presets(loader, para_text) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

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
        cursor_p = Write.get_paragraph_cursor(cursor)
        Write.append_para(cursor=cursor, text="Testing Presets.")
        for preset in presets:
            hatch = Hatch.from_preset(preset)
            Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
            cursor_p.gotoPreviousParagraph(False)
            cursor_p.gotoStartOfParagraph(False)
            cursor_p.gotoEndOfParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            assert fp.FillStyle == FillStyle.HATCH
            assert fp.FillBackground == False
            assert fp.FillColor == StandardColor.DEFAULT_BLUE
            assert fp.FillHatchName == str(preset)
            pp = cast("ParagraphProperties", cursor_p)
            assert pp.ParaBackTransparent == False
            cursor_p.gotoEnd(False)

        Write.append_para(cursor=cursor, text="Testing Presets With Backcolor.")
        for preset in presets:
            bg_color = StandardColor.get_random_color()
            hatch = Hatch.from_preset(preset)
            hatch.prop_bg_color = bg_color
            Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
            cursor_p.gotoPreviousParagraph(False)
            cursor_p.gotoStartOfParagraph(False)
            cursor_p.gotoEndOfParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            assert fp.FillStyle == FillStyle.HATCH
            assert fp.FillBackground == True
            assert fp.FillColor == bg_color
            assert fp.FillHatchName == str(preset)
            pp = cast("ParagraphProperties", cursor_p)
            assert pp.ParaBackTransparent == False
            cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write(loader, para_text) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        cursor_p = Write.get_paragraph_cursor(cursor)

        hatch = Hatch.from_preset(PresetHatchKind.GREEN_30_DEGREES)
        Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == False
        assert fp.FillColor == StandardColor.DEFAULT_BLUE
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackTransparent == False
        cursor_p.gotoEnd(False)

        hatch = Hatch.from_preset(PresetHatchKind.GREEN_30_DEGREES)
        hatch.prop_color = StandardColor.BLUE_LIGHT2
        Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == False
        assert fp.FillColor == StandardColor.DEFAULT_BLUE
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackTransparent == False
        cursor_p.gotoEnd(False)

        hatch = Hatch.from_preset(PresetHatchKind.GREEN_30_DEGREES)
        hatch.prop_space = 2.7
        hatch.prop_color = StandardColor.BRICK_DARK3
        hatch.prop_bg_color = StandardColor.BRICK_LIGHT2
        Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == True
        assert fp.FillColor == StandardColor.BRICK_LIGHT2
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackTransparent == False
        cursor_p.gotoEnd(False)

        hatch = Hatch(style=HatchStyle.DOUBLE, color=StandardColor.BRICK_LIGHT1, space=2.1)
        Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == False
        assert fp.FillColor == StandardColor.DEFAULT_BLUE
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackTransparent == False
        cursor_p.gotoEnd(False)

        hatch = Hatch(
            style=HatchStyle.DOUBLE, color=StandardColor.GOLD_DARK2, space=2.1, name="Custom Hatch", auto_name=False
        )
        Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == False
        assert fp.FillColor == StandardColor.DEFAULT_BLUE
        assert fp.FillHatchName == "Custom Hatch"
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackTransparent == False
        cursor_p.gotoEnd(False)

        hatch = Hatch(
            style=HatchStyle.SINGLE, color=StandardColor.GREEN_DARK2, space=2.1, name="Custom Hatch", auto_name=False
        )
        Write.append_para(cursor=cursor, text=para_text, styles=(hatch,))
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.HATCH
        assert fp.FillBackground == False
        assert fp.FillColor == StandardColor.DEFAULT_BLUE
        assert fp.FillHatchName == "Custom Hatch"
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackTransparent == False

        cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
