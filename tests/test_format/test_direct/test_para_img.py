from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.format.direct.para.area.img import (
    Img,
    PresetImageKind,
    ImgStyleKind,
    SizeMM,
    SizePercent,
    Offset,
    OffsetColumn,
    OffsetRow,
    RectanglePoint,
)
from ooodev.utils.images_lo import ImagesLo, BitmapArgs

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
            PresetImageKind.PAINTED_WHITE,
            PresetImageKind.PAPER_TEXTURE,
            PresetImageKind.PAPER_CRUMPLED,
            PresetImageKind.PARCHMENT_PAPER,
            PresetImageKind.FENCE,
            PresetImageKind.INVOICE_PAPER,
            PresetImageKind.MAPLE_LEAVES,
            PresetImageKind.LAWN,
            PresetImageKind.COLORFUL_PEBBLES,
            PresetImageKind.COFFEE_BEANS,
            PresetImageKind.LITTLE_CLOUDS,
            PresetImageKind.BATHROOM_TILES,
            PresetImageKind.CONCRETE,
            PresetImageKind.PAPER_GRAPH,
            PresetImageKind.ZEBRA,
            PresetImageKind.WALL_OF_ROCK,
            PresetImageKind.BRICK_WALL,
            PresetImageKind.STONE_WALL,
            PresetImageKind.FLORAL,
            PresetImageKind.SPACE,
            PresetImageKind.COLOR_STRIPES,
            PresetImageKind.ICE_LIGHT,
            PresetImageKind.POOL,
            PresetImageKind.MARBLE,
            PresetImageKind.SAND_LIGHT,
            PresetImageKind.STONE,
            PresetImageKind.WHITE_DIFFUSION,
            PresetImageKind.SURFACE,
            PresetImageKind.CARDBOARD,
            PresetImageKind.STUDIO,
        )
        cursor_p = Write.get_paragraph_cursor(cursor)
        for preset in presets:
            img = Img.from_preset(preset)

            Write.append_para(cursor=cursor, text=para_text, styles=(img,))

            cursor_p.gotoPreviousParagraph(False)
            cursor_p.gotoStartOfParagraph(False)
            cursor_p.gotoEndOfParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            # point = preset._get_point()
            # assert fp.FillBitmapSizeX == point.x
            # assert fp.FillBitmapSizeY == point.y
            assert fp.FillBitmapName == str(preset)
            assert fp.FillBitmap is not None
            assert fp.FillBitmapMode == ImgStyleKind.TILED.value
            pp = cast("ParagraphProperties", cursor_p)
            assert pp.ParaBackColor == -1
            assert pp.ParaBackTransparent == True

            cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_image(loader, para_text, fix_image_path) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        fnm = fix_image_path("img_brick.png")
        bargs = BitmapArgs("Bitmap ", auto_name=True)
        xb = ImagesLo.get_bitmap(fnm, bargs)
        img = Img(
            bitmap=xb,
            name=bargs.out_name,
            mode=ImgStyleKind.TILED,
            size=SizeMM(67.73, 67.73),
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
        )

        cursor = Write.get_cursor(doc)
        cursor_p = Write.get_paragraph_cursor(cursor)

        Write.append_para(cursor=cursor, text=para_text, styles=(img,))

        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillBitmapName == bargs.out_name
        assert fp.FillBitmap is not None
        assert fp.FillBitmapMode == ImgStyleKind.TILED.value
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackTransparent == True

        cursor_p.gotoEnd(False)

        # Test again this time no bitmap just name
        img = Img(
            name=bargs.out_name,
            mode=ImgStyleKind.TILED,
            size=SizeMM(67.73, 67.73),
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
        )

        Write.append_para(cursor=cursor, text=para_text, styles=(img,))

        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillBitmapName == bargs.out_name
        assert fp.FillBitmap is not None
        assert fp.FillBitmapMode == ImgStyleKind.TILED.value
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackTransparent == True

        cursor_p.gotoEnd(False)

        Lo.delay(delay)

    finally:
        Lo.close_doc(doc)
