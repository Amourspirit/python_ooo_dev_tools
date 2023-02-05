from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.format.direct.para.area.pattern import Pattern, PatternKind
from ooodev.utils.images_lo import ImagesLo, BitmapArgs


from ooo.dyn.style.graphic_location import GraphicLocation

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        pattern = Pattern.from_preset(PatternKind.DIVOT)
        Write.append_para(cursor=cursor, text=para_text, styles=(pattern,))

        cursor_p = Write.get_paragraph_cursor(cursor)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        # note: it is necessary to reast fp each time cursor_p is moved
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == str(PatternKind.DIVOT)
        assert fp.FillBitmap is not None
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackGraphicLocation == GraphicLocation.TILED
        assert pp.ParaBackTransparent == True
        assert pp.ParaBackGraphic is not None
        cursor_p.gotoEnd(False)

        pattern = Pattern.from_preset(PatternKind.DIAGONAL_BRICK)
        Write.append_para(cursor=cursor, text=para_text, styles=(pattern,))

        cursor_p = Write.get_paragraph_cursor(cursor)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        # note: it is necessary to reast fp each time cursor_p is moved
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == str(PatternKind.DIAGONAL_BRICK)
        assert fp.FillBitmap is not None
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackGraphicLocation == GraphicLocation.TILED
        assert pp.ParaBackTransparent == True
        assert pp.ParaBackGraphic is not None
        cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_from_image(loader, para_text, fix_image_path) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        fnm = fix_image_path("pattern_brick.png")
        bargs = BitmapArgs("Bitmap ", auto_name=True)
        xb = ImagesLo.get_bitmap(fnm, bargs)

        cursor = Write.get_cursor(doc)

        pattern = Pattern(bitmap=xb, name=bargs.out_name)
        Write.append_para(cursor=cursor, text=para_text, styles=(pattern,))

        cursor_p = Write.get_paragraph_cursor(cursor)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        # note: it is necessary to reast fp each time cursor_p is moved
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == bargs.out_name
        assert fp.FillBitmap is not None
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackGraphicLocation == GraphicLocation.TILED
        assert pp.ParaBackTransparent == True
        assert pp.ParaBackGraphic is not None
        cursor_p.gotoEnd(False)

        # test no bitmap just name on second pass, This should work on first pass alos
        # because call to ImagesLo.get_bitmap() would have added image to Bitmap Table.
        pattern = Pattern(name=bargs.out_name)
        Write.append_para(cursor=cursor, text=para_text, styles=(pattern,))

        cursor_p = Write.get_paragraph_cursor(cursor)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        # note: it is necessary to reast fp each time cursor_p is moved
        assert fp.FillBitmapTile == True
        assert fp.FillBitmapStretch == False
        assert fp.FillBitmapName == bargs.out_name
        assert fp.FillBitmap is not None
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackGraphicLocation == GraphicLocation.TILED
        assert pp.ParaBackTransparent == True
        assert pp.ParaBackGraphic is not None
        cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
